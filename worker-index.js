// @ts-nocheck
/**
 * GIF Local — API de licenciamento
 * Cloudflare Worker + D1. Assina licenças com ECDSA P-256 (Web Crypto).
 * A chave privada NUNCA fica neste arquivo — vem de env.LICENSE_PRIVATE_KEY_JWK (secret).
 */

const ALLOWED_ORIGIN = 'https://giflocal.pages.dev';

// Nomes batem com os planos já existentes nos Payment Links do Stripe
// (ver worker/README.md, passo 6, sobre como marcar metadata.plan em cada link).
const PLAN_DEVICES = {
  apoiador: 1,
  profissional: 3,
  empresa: 10, // plano ainda sem Payment Link público — crie a licença via wrangler d1 execute se negociar diretamente
};

const OFFLINE_TOLERANCE_DAYS = { default: 30, lifetime: 90 };

function corsHeaders(extra = {}) {
  return {
    'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
    'Access-Control-Allow-Methods': 'GET,POST,DELETE,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization',
    'Content-Type': 'application/json; charset=utf-8',
    ...extra,
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: corsHeaders() });
}

function err(code, message, status = 400) {
  return json({ error: code, message }, status);
}

function nowIso() {
  return new Date().toISOString();
}

function uuid() {
  return crypto.randomUUID();
}

const KEY_ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // sem 0/O/1/I para evitar confusão ao digitar

function generateLicenseKey() {
  const groups = [];
  for (let g = 0; g < 4; g++) {
    let group = '';
    const bytes = crypto.getRandomValues(new Uint8Array(4));
    for (let i = 0; i < 4; i++) group += KEY_ALPHABET[bytes[i] % KEY_ALPHABET.length];
    groups.push(group);
  }
  return 'GLPR-' + groups.join('-');
}

async function sha256Hex(input) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(input));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function base64FromBuf(buf) {
  let bin = '';
  const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

async function getSigningKey(env) {
  const jwk = JSON.parse(env.LICENSE_PRIVATE_KEY_JWK);
  return crypto.subtle.importKey('jwk', jwk, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign']);
}

/** Assina um payload de licença. Retorna { payload: <string JSON>, signature: <base64> }. */
async function signLicensePayload(env, payloadObj) {
  const payloadString = JSON.stringify(payloadObj);
  const key = await getSigningKey(env);
  const sigBuf = await crypto.subtle.sign(
    { name: 'ECDSA', hash: 'SHA-256' },
    key,
    new TextEncoder().encode(payloadString)
  );
  return { payload: payloadString, signature: base64FromBuf(sigBuf) };
}

function offlineToleranceDays(license) {
  return license.expires_at ? OFFLINE_TOLERANCE_DAYS.default : OFFLINE_TOLERANCE_DAYS.lifetime;
}

function effectiveStatus(license) {
  if (license.status === 'REVOKED' || license.status === 'SUSPENDED') return license.status;
  if (license.expires_at && new Date(license.expires_at).getTime() < Date.now()) return 'EXPIRED';
  return license.status;
}

async function buildSignedLicense(env, license, device) {
  const status = effectiveStatus(license);
  const activeDevices = await countActiveDevices(env, license.license_id);
  const payloadObj = {
    licenseId: license.license_id,
    deviceId: device ? device.device_id : null,
    plan: license.plan,
    status,
    maxDevices: license.max_devices,
    activeDevices,
    expiresAt: license.expires_at,
    offlineToleranceDays: offlineToleranceDays(license),
    issuedAt: nowIso(),
  };
  return signLicensePayload(env, payloadObj);
}

async function countActiveDevices(env, licenseId) {
  const row = await env.DB.prepare(
    'SELECT COUNT(*) AS n FROM devices WHERE license_id = ? AND deactivated_at IS NULL'
  )
    .bind(licenseId)
    .first();
  return row ? row.n : 0;
}

async function logEvent(env, { licenseId, deviceId, event, detail, request, appVersion }) {
  const ip = request ? request.headers.get('CF-Connecting-IP') || '' : '';
  const ipHash = ip ? await sha256Hex(ip) : null;
  await env.DB.prepare(
    `INSERT INTO activation_log (license_id, device_id, event, detail, ip_hash, app_version, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?)`
  )
    .bind(licenseId || null, deviceId || null, event, detail || null, ipHash, appVersion || null, nowIso())
    .run();
}

async function readJson(request) {
  try {
    return await request.json();
  } catch {
    return null;
  }
}

/* ============================== rotas ============================== */

async function handleActivate(request, env) {
  const body = await readJson(request);
  if (!body || !body.licenseKey || !body.deviceId) return err('BAD_REQUEST', 'licenseKey e deviceId são obrigatórios');

  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_key = ?')
    .bind(body.licenseKey.trim().toUpperCase())
    .first();
  if (!license) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);

  const status = effectiveStatus(license);
  if (status === 'REVOKED' || status === 'SUSPENDED' || status === 'EXPIRED') {
    return err('LICENSE_NOT_ACTIVATABLE', `Licença está ${status}`, 403);
  }

  let device = await env.DB.prepare('SELECT * FROM devices WHERE device_id = ? AND license_id = ?')
    .bind(body.deviceId, license.license_id)
    .first();

  if (device && device.deactivated_at) {
    // dispositivo foi desativado antes; reativar consome uma vaga de novo
    device = null;
  }

  if (!device) {
    const activeCount = await countActiveDevices(env, license.license_id);
    if (activeCount >= license.max_devices) {
      return err(
        'LICENSE_DEVICE_LIMIT_REACHED',
        'Esta licença já está ativada no número máximo de dispositivos permitido.',
        409
      );
    }
    await env.DB.prepare(
      `INSERT INTO devices (device_id, license_id, device_label, first_seen_at, last_validated_at, deactivated_at)
       VALUES (?, ?, ?, ?, ?, NULL)`
    )
      .bind(body.deviceId, license.license_id, body.deviceLabel || null, nowIso(), nowIso())
      .run();
  } else {
    await env.DB.prepare('UPDATE devices SET last_validated_at = ? WHERE device_id = ?')
      .bind(nowIso(), body.deviceId)
      .run();
  }

  if (license.status === 'PENDING') {
    await env.DB.prepare('UPDATE licenses SET status = ?, activated_at = ? WHERE license_id = ?')
      .bind('ACTIVE', nowIso(), license.license_id)
      .run();
    license.status = 'ACTIVE';
    license.activated_at = nowIso();
  }

  const signed = await buildSignedLicense(env, license, { device_id: body.deviceId });
  await logEvent(env, { licenseId: license.license_id, deviceId: body.deviceId, event: 'ACTIVATE', request, appVersion: body.appVersion });
  return json(signed);
}

async function handleValidate(request, env) {
  const body = await readJson(request);
  if (!body || !body.licenseId || !body.deviceId) return err('BAD_REQUEST', 'licenseId e deviceId são obrigatórios');

  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_id = ?').bind(body.licenseId).first();
  if (!license) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);

  const device = await env.DB.prepare('SELECT * FROM devices WHERE device_id = ? AND license_id = ?')
    .bind(body.deviceId, body.licenseId)
    .first();
  if (!device) return err('DEVICE_NOT_FOUND', 'Dispositivo não registrado nesta licença', 404);

  if (device.deactivated_at) {
    // dispositivo foi desativado (ex.: pelo fluxo de "gerenciar dispositivos") — assina um
    // payload dizendo isso, em vez de simplesmente recusar, para o app conseguir mostrar
    // uma mensagem clara e derrubar o estado PRO local de forma confiável.
    const payloadObj = {
      licenseId: license.license_id,
      deviceId: body.deviceId,
      plan: license.plan,
      status: 'DEVICE_DEACTIVATED',
      maxDevices: license.max_devices,
      activeDevices: await countActiveDevices(env, license.license_id),
      expiresAt: license.expires_at,
      offlineToleranceDays: offlineToleranceDays(license),
      issuedAt: nowIso(),
    };
    const signed = await signLicensePayload(env, payloadObj);
    await logEvent(env, { licenseId: license.license_id, deviceId: body.deviceId, event: 'VALIDATE', detail: 'device_deactivated', request });
    return json(signed);
  }

  await env.DB.prepare('UPDATE devices SET last_validated_at = ? WHERE device_id = ?').bind(nowIso(), body.deviceId).run();
  const signed = await buildSignedLicense(env, license, device);
  await logEvent(env, { licenseId: license.license_id, deviceId: body.deviceId, event: 'VALIDATE', detail: effectiveStatus(license), request });
  return json(signed);
}

async function handleListDevices(request, env) {
  const url = new URL(request.url);
  const licenseKey = (url.searchParams.get('licenseKey') || '').trim().toUpperCase();
  if (!licenseKey) return err('BAD_REQUEST', 'licenseKey é obrigatório');

  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_key = ?').bind(licenseKey).first();
  if (!license) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);

  const { results } = await env.DB.prepare(
    `SELECT device_id, device_label, first_seen_at, last_validated_at, deactivated_at
     FROM devices WHERE license_id = ? ORDER BY first_seen_at DESC`
  )
    .bind(license.license_id)
    .all();

  return json({ licenseId: license.license_id, plan: license.plan, maxDevices: license.max_devices, devices: results });
}

async function handleDeactivateDevice(request, env) {
  const body = await readJson(request);
  if (!body || !body.licenseKey || !body.deviceId) return err('BAD_REQUEST', 'licenseKey e deviceId são obrigatórios');

  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_key = ?')
    .bind(body.licenseKey.trim().toUpperCase())
    .first();
  if (!license) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);

  const device = await env.DB.prepare('SELECT * FROM devices WHERE device_id = ? AND license_id = ?')
    .bind(body.deviceId, license.license_id)
    .first();
  if (!device) return err('DEVICE_NOT_FOUND', 'Dispositivo não encontrado nesta licença', 404);

  await env.DB.prepare('UPDATE devices SET deactivated_at = ? WHERE device_id = ?').bind(nowIso(), body.deviceId).run();
  await logEvent(env, { licenseId: license.license_id, deviceId: body.deviceId, event: 'DEACTIVATE', request });
  return json({ ok: true });
}

async function handleLicenseBySession(request, env) {
  const url = new URL(request.url);
  const sessionId = url.searchParams.get('session_id');
  if (!sessionId) return err('BAD_REQUEST', 'session_id é obrigatório');

  const license = await env.DB.prepare('SELECT license_key, plan, status FROM licenses WHERE stripe_checkout_session_id = ?')
    .bind(sessionId)
    .first();
  if (!license) return err('NOT_READY', 'Ainda processando o pagamento — tente novamente em alguns segundos.', 404);

  return json(license);
}

/* ---- Stripe webhook ---- */

async function verifyStripeSignature(request, env, rawBody) {
  const sigHeader = request.headers.get('Stripe-Signature') || '';
  const parts = Object.fromEntries(
    sigHeader.split(',').map((kv) => {
      const [k, v] = kv.split('=');
      return [k, v];
    })
  );
  if (!parts.t || !parts.v1) return false;

  const signedPayload = `${parts.t}.${rawBody}`;
  const key = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(env.STRIPE_WEBHOOK_SECRET),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const sigBuf = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(signedPayload));
  const expectedHex = [...new Uint8Array(sigBuf)].map((b) => b.toString(16).padStart(2, '0')).join('');
  return expectedHex === parts.v1;
}

async function handleStripeWebhook(request, env) {
  const rawBody = await request.text();
  const validSig = await verifyStripeSignature(request, env, rawBody);
  if (!validSig) return err('INVALID_SIGNATURE', 'Assinatura do webhook inválida', 400);

  const event = JSON.parse(rawBody);

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object;
    const email = session.customer_details?.email || session.customer_email;
    const plan = (session.metadata && session.metadata.plan) || null;
    if (!email || !plan || !PLAN_DEVICES[plan]) {
      await logEvent(env, { event: 'WEBHOOK', detail: `checkout.session.completed sem email/plan válido (session ${session.id})`, request });
      return json({ ok: true, warning: 'plan/email ausente — verifique metadata do Payment Link' });
    }

    let customer = await env.DB.prepare('SELECT * FROM customers WHERE email = ?').bind(email).first();
    if (!customer) {
      const customerId = uuid();
      await env.DB.prepare('INSERT INTO customers (customer_id, email, stripe_customer_id, created_at) VALUES (?, ?, ?, ?)')
        .bind(customerId, email, session.customer || null, nowIso())
        .run();
      customer = { customer_id: customerId };
    }

    const licenseId = uuid();
    const licenseKey = generateLicenseKey();
    await env.DB.prepare(
      `INSERT INTO licenses
        (license_id, license_key, customer_id, product, plan, status, max_devices,
         stripe_checkout_session_id, stripe_payment_id, purchased_at)
       VALUES (?, ?, ?, 'giflocal', ?, 'PENDING', ?, ?, ?, ?)`
    )
      .bind(licenseId, licenseKey, customer.customer_id, plan, PLAN_DEVICES[plan], session.id, session.payment_intent || null, nowIso())
      .run();

    await logEvent(env, { licenseId, event: 'WEBHOOK', detail: 'checkout.session.completed -> licença criada', request });
  } else if (event.type === 'charge.refunded' || event.type === 'charge.dispute.created') {
    const obj = event.data.object;
    const paymentIntent = obj.payment_intent;
    if (paymentIntent) {
      const license = await env.DB.prepare('SELECT * FROM licenses WHERE stripe_payment_id = ?').bind(paymentIntent).first();
      if (license) {
        const newStatus = event.type === 'charge.refunded' ? 'REVOKED' : 'SUSPENDED';
        const reason = event.type === 'charge.refunded' ? 'refund' : 'dispute';
        await env.DB.prepare('UPDATE licenses SET status = ?, revoked_at = ?, revoked_reason = ? WHERE license_id = ?')
          .bind(newStatus, nowIso(), reason, license.license_id)
          .run();
        await logEvent(env, { licenseId: license.license_id, event: 'WEBHOOK', detail: `${event.type} -> ${newStatus}`, request });
      }
    }
  }

  return json({ ok: true });
}

/* ---- admin leve (protegido por ADMIN_TOKEN), sem interface — só endpoints ---- */

function requireAdmin(request, env) {
  const auth = request.headers.get('Authorization') || '';
  return auth === `Bearer ${env.ADMIN_TOKEN}`;
}

async function handleAdminGetLicense(request, env, licenseId) {
  if (!requireAdmin(request, env)) return err('UNAUTHORIZED', 'Token inválido', 401);
  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_id = ?').bind(licenseId).first();
  if (!license) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);
  const { results: devices } = await env.DB.prepare('SELECT * FROM devices WHERE license_id = ?').bind(licenseId).all();
  return json({ license, devices });
}

async function handleAdminSetStatus(request, env, licenseId) {
  if (!requireAdmin(request, env)) return err('UNAUTHORIZED', 'Token inválido', 401);
  const body = await readJson(request);
  const allowed = ['ACTIVE', 'REVOKED', 'SUSPENDED', 'EXPIRED'];
  if (!body || !allowed.includes(body.status)) return err('BAD_REQUEST', `status deve ser um de: ${allowed.join(', ')}`);
  await env.DB.prepare('UPDATE licenses SET status = ?, revoked_at = ?, revoked_reason = ? WHERE license_id = ?')
    .bind(body.status, body.status === 'ACTIVE' ? null : nowIso(), body.reason || null, licenseId)
    .run();
  await logEvent(env, { licenseId, event: 'REVOKE', detail: `admin -> ${body.status} (${body.reason || ''})`, request });
  return json({ ok: true });
}

/* ====================== geração de imagem com IA ======================
   A IA gera uma IMAGEM. O movimento, a otimização e o GIF continuam sendo
   feitos no navegador do usuário — isso mantém a promessa de privacidade
   honesta (só o texto do prompt sai do dispositivo, nunca os arquivos) e
   derruba o custo, porque não existe geração de vídeo gratuita.

   Usa Workers AI pelo binding `env.AI` — não há API key envolvida.

   Dois limites, com propósitos diferentes:
   - AI_DAILY_LIMIT: quota comercial por usuário (é o que vira paywall)
   - AI_GLOBAL_DAILY_CAP: protege a cota diária GRATUITA da conta inteira
     (10.000 neurons/dia). Sem esse teto, um punhado de usuários esgotaria
     a franquia e derrubaria a feature para todo mundo — ou geraria custo
     inesperado. É uma trava de segurança financeira, não comercial.
   ===================================================================== */

const AI_MODEL = '@cf/black-forest-labs/flux-1-schnell';
const AI_STEPS = 4; /* schnell é treinado para 4 passos; mais que isso gasta neurons sem ganho */

const AI_DAILY_LIMIT = { free: 1, pro: 25, full: 80 };

/* MEDIDO em produção (painel Workers AI): 345,6 neurons para 2 imagens =
   172,8 por imagem. O flux-1-schnell gera 1024x1024 fixo — são 4 tiles de
   512x512, e o custo dos passos é por tile, não por imagem:
     4 tiles x 4,80  +  4 passos x 4 tiles x 9,60  =  172,8
   O modelo não aceita width/height, então não dá para pedir menor.
   10.000 neurons/dia gratuitos ÷ 172,8 = 57 imagens/dia. Fixamos 55 para
   deixar margem — passar disso não quebra nada, mas sai do gratuito. */
const AI_GLOBAL_DAILY_CAP = 55;

const BACKEND_PLAN_TO_TIER = { apoiador: 'pro', profissional: 'full', empresa: 'full' };

function todayKey() {
  return nowIso().slice(0, 10);
}

/* Descobre o tier de quem está pedindo. Sem licença = free. A chave de
   licença é conferida no banco (nunca confiamos num "plano" enviado pelo
   navegador, que qualquer um poderia forjar). */
async function resolveAiSubject(env, body) {
  const deviceId = String((body && body.deviceId) || '').slice(0, 120);
  let tier = 'free';
  let subject = 'dev:' + (deviceId || 'anon');
  const key = body && body.licenseKey ? String(body.licenseKey).trim().toUpperCase() : '';
  if (key) {
    const lic = await env.DB.prepare('SELECT license_id, plan, status FROM licenses WHERE license_key = ?')
      .bind(key)
      .first();
    if (lic && lic.status === 'ACTIVE') {
      tier = BACKEND_PLAN_TO_TIER[lic.plan] || 'pro';
      subject = 'lic:' + lic.license_id;
    }
  }
  return { tier, subject };
}

async function readAiUsage(env, subject) {
  const day = todayKey();
  const mine = await env.DB.prepare('SELECT count FROM ai_usage WHERE day = ? AND subject = ?')
    .bind(day, subject)
    .first();
  const total = await env.DB.prepare('SELECT COALESCE(SUM(count), 0) AS n FROM ai_usage WHERE day = ?')
    .bind(day)
    .first();
  return { day, used: (mine && mine.count) || 0, globalUsed: (total && total.n) || 0 };
}

async function handleAiQuota(request, env) {
  const url = new URL(request.url);
  const body = { deviceId: url.searchParams.get('deviceId'), licenseKey: url.searchParams.get('licenseKey') };
  const { tier, subject } = await resolveAiSubject(env, body);
  const { used, globalUsed } = await readAiUsage(env, subject);
  const allowed = AI_DAILY_LIMIT[tier];
  return json({
    tier,
    used,
    allowed,
    remaining: Math.max(0, allowed - used),
    serviceBusy: globalUsed >= AI_GLOBAL_DAILY_CAP,
  });
}

async function handleAiGenerate(request, env) {
  if (!env.AI) return err('AI_NOT_CONFIGURED', 'Geração com IA ainda não está ativada neste servidor.', 503);

  const body = await readJson(request);
  const prompt = body && typeof body.prompt === 'string' ? body.prompt.trim() : '';
  if (!prompt) return err('BAD_REQUEST', 'Descreva o que você quer criar.');
  if (prompt.length > 800) return err('PROMPT_TOO_LONG', 'Descrição muito longa (máximo 800 caracteres).');

  const { tier, subject } = await resolveAiSubject(env, body);
  const { day, used, globalUsed } = await readAiUsage(env, subject);
  const allowed = AI_DAILY_LIMIT[tier];

  if (globalUsed >= AI_GLOBAL_DAILY_CAP) {
    return json(
      { error: 'AI_SERVICE_BUSY', message: 'O limite diário de gerações com IA do serviço foi atingido. Tente de novo amanhã.' },
      429
    );
  }
  if (used >= allowed) {
    return json(
      { error: 'AI_LIMIT_REACHED', message: 'Você usou suas gerações com IA de hoje.', tier, used, allowed },
      429
    );
  }

  /* Conta ANTES de gerar: se contássemos depois, chamadas simultâneas
     furariam o teto e poderiam estourar a franquia da conta. O custo de
     uma geração que falhar é um crédito perdido — preferível a um limite
     que não segura. */
  await env.DB.prepare(
    `INSERT INTO ai_usage (day, subject, count) VALUES (?, ?, 1)
     ON CONFLICT(day, subject) DO UPDATE SET count = count + 1`
  )
    .bind(day, subject)
    .run();

  const result = await env.AI.run(AI_MODEL, { prompt, steps: AI_STEPS });

  /* O formato de retorno varia entre modelos: alguns devolvem { image: base64 },
     outros um ReadableStream binário. Normalizamos para base64 aqui, para o
     frontend não precisar saber qual modelo está por trás. */
  let imageB64 = null;
  if (result && typeof result.image === 'string') {
    imageB64 = result.image;
  } else if (result instanceof ReadableStream || (result && typeof result.getReader === 'function')) {
    const buf = await new Response(result).arrayBuffer();
    imageB64 = arrayBufferToBase64(buf);
  } else if (result instanceof ArrayBuffer) {
    imageB64 = arrayBufferToBase64(result);
  }

  if (!imageB64) {
    return err('AI_UNEXPECTED_OUTPUT', 'O modelo respondeu num formato inesperado: ' + typeof result, 502);
  }

  return json({
    ok: true,
    image: imageB64,
    model: AI_MODEL,
    tier,
    used: used + 1,
    allowed,
    remaining: Math.max(0, allowed - used - 1),
  });
}

function arrayBufferToBase64(buf) {
  const bytes = new Uint8Array(buf);
  let bin = '';
  const CHUNK = 0x8000; /* converter de uma vez estoura a pilha em imagens grandes */
  for (let i = 0; i < bytes.length; i += CHUNK) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return btoa(bin);
}

/* ============================== router ============================== */

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders() });
    }

    try {
      if (url.pathname === '/activate' && request.method === 'POST') return await handleActivate(request, env);
      if (url.pathname === '/validate' && request.method === 'POST') return await handleValidate(request, env);
      if (url.pathname === '/license/devices' && request.method === 'GET') return await handleListDevices(request, env);
      if (url.pathname === '/device/deactivate' && request.method === 'POST') return await handleDeactivateDevice(request, env);
      if (url.pathname === '/license-by-session' && request.method === 'GET') return await handleLicenseBySession(request, env);
      if (url.pathname === '/webhooks/stripe' && request.method === 'POST') return await handleStripeWebhook(request, env);
      if (url.pathname === '/ai/generate' && request.method === 'POST') return await handleAiGenerate(request, env);
      if (url.pathname === '/ai/quota' && request.method === 'GET') return await handleAiQuota(request, env);

      const adminMatch = url.pathname.match(/^\/admin\/license\/([^/]+)(\/status)?$/);
      if (adminMatch && request.method === 'GET' && !adminMatch[2]) return await handleAdminGetLicense(request, env, adminMatch[1]);
      if (adminMatch && request.method === 'POST' && adminMatch[2] === '/status')
        return await handleAdminSetStatus(request, env, adminMatch[1]);

      return err('NOT_FOUND', 'Rota não encontrada', 404);
    } catch (e) {
      return err('INTERNAL_ERROR', String(e && e.message ? e.message : e), 500);
    }
  },
};
