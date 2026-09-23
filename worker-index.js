// @ts-nocheck
/**
 * GIF Local — API de licenciamento
 * Cloudflare Worker + D1. Assina licenças com ECDSA P-256 (Web Crypto).
 * A chave privada NUNCA fica neste arquivo — vem de env.LICENSE_PRIVATE_KEY_JWK (secret).
 */

/* O site é servido em DOIS domínios: giflocal.com (o domínio de verdade,
   para onde os Payment Links do Stripe mandam o cliente) e
   giflocal.pages.dev (a URL do Cloudflare Pages). Enquanto aqui só
   constava o pages.dev, TUDO que passa por esta API estava quebrado em
   giflocal.com — ativar licença, gerenciar dispositivos, gerar com IA —
   com um "Failed to fetch" seco no navegador, porque o CORS recusava a
   resposta. A requisição chegava ao servidor; era a resposta que o
   navegador jogava fora.

   ALLOWED_ORIGIN continua sendo o padrão para quem chega sem cabeçalho
   Origin (curl, webhook do Stripe). */
const ALLOWED_ORIGIN = 'https://giflocal.pages.dev';
const ALLOWED_ORIGINS = [
  'https://giflocal.com',
  'https://www.giflocal.com',
  'https://giflocal.pages.dev',
];

function resolveOrigin(request) {
  const origin = request && request.headers ? request.headers.get('Origin') : null;
  return origin && ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGIN;
}

/* Corrige o Access-Control-Allow-Origin de uma resposta já pronta.

   Feito num lugar só, na saída do roteador, de propósito: a alternativa
   seria passar o request por json()/err() e por todos os handlers. Um
   "origem atual" em variável de módulo seria mais curto e ERRADO — no
   Workers o mesmo isolate atende requisições concorrentes, e duas
   chamadas de domínios diferentes poderiam trocar de origem uma com a
   outra. */
function withCors(response, request) {
  const headers = new Headers(response.headers);
  headers.set('Access-Control-Allow-Origin', resolveOrigin(request));
  headers.append('Vary', 'Origin'); // senão um cache serviria a resposta de um domínio para o outro
  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

/* ============================ segurança ============================
   Utilitários usados pelas rotas. Motivo de cada um está no relatório de
   auditoria (set/2026); o resumo vai junto de cada função. */

/* Aceita só texto, aparado, dentro do limite. Qualquer outro tipo (número,
   objeto, lista) vira null: antes, { "licenseKey": 123 } derrubava a rota
   com um 500 que devolvia a mensagem interna do JavaScript. */
function texto(v, max) {
  if (typeof v !== 'string') return null;
  const t = v.trim();
  if (!t || t.length > max) return null;
  return t;
}

/* Comparação em tempo constante: não revela, pelo tempo de resposta,
   quantos caracteres iniciais de um segredo o atacante já acertou. */
function iguaisTempoConstante(a, b) {
  const ea = new TextEncoder().encode(String(a));
  const eb = new TextEncoder().encode(String(b));
  if (ea.length !== eb.length) return false;
  let dif = 0;
  for (let i = 0; i < ea.length; i++) dif |= ea[i] ^ eb[i];
  return dif === 0;
}

/* Em IPv6, um único assinante recebe um bloco /64 inteiro (18 quintilhões
   de endereços): limitar por endereço exato não limitaria nada. Agrupamos
   pelos 4 primeiros grupos, que identificam a conexão. */
function prefixoIp(ip) {
  if (!ip || ip.indexOf(':') < 0) return ip;
  let [cabeca, cauda] = ip.split('::');
  const a = cabeca ? cabeca.split(':') : [];
  const b = cauda !== undefined && cauda ? cauda.split(':') : [];
  const faltam = cauda !== undefined ? 8 - a.length - b.length : 0;
  const grupos = [...a, ...Array(Math.max(0, faltam)).fill('0'), ...b];
  return grupos.slice(0, 4).map((g) => (parseInt(g, 16) || 0).toString(16)).join(':') + '::/64';
}

/* Hash do IP com chave. SHA-256 puro de um IPv4 se desfaz por força bruta
   (são só 4 bilhões de possibilidades); com HMAC e uma chave que só o
   servidor conhece, o hash continua útil para limitar abuso e deixa de ser
   um dado pessoal recuperável. A chave deriva de um segredo que já existe,
   sem criar configuração nova. */
async function hashIp(env, ipBruto) {
  const ip = prefixoIp(ipBruto);
  if (!ip) return null;
  const base = new TextEncoder().encode('ip-hash:' + (env.LICENSE_PRIVATE_KEY_JWK || ''));
  const chave = await crypto.subtle.importKey('raw', await crypto.subtle.digest('SHA-256', base), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const sig = await crypto.subtle.sign('HMAC', chave, new TextEncoder().encode(ip));
  return [...new Uint8Array(sig)].slice(0, 16).map((b) => b.toString(16).padStart(2, '0')).join('');
}

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
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization',
    'Access-Control-Max-Age': '600',
    'Content-Type': 'application/json; charset=utf-8',
    /* Respostas carregam chave de licença: nenhum cache intermediário pode
       guardar, e o navegador não deve reinterpretar o tipo. */
    'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'no-referrer',
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
  const ipHash = await hashIp(env, ip);
  await env.DB.prepare(
    `INSERT INTO activation_log (license_id, device_id, event, detail, ip_hash, app_version, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?)`
  )
    .bind(licenseId || null, deviceId || null, event, detail || null, ipHash, appVersion || null, nowIso())
    .run();
}

/* Lê o corpo com teto de tamanho contando os bytes que chegam, e não só
   o cabeçalho Content-Length (um corpo "chunked" não tem esse cabeçalho).
   Devolve null se passar do teto. */
async function leCorpo(request, limite) {
  if (!request.body) return '';
  const leitor = request.body.getReader();
  const partes = [];
  let total = 0;
  for (;;) {
    const { done, value } = await leitor.read();
    if (done) break;
    total += value.byteLength;
    if (total > limite) { try { await leitor.cancel(); } catch {} return null; }
    partes.push(value);
  }
  const buf = new Uint8Array(total);
  let p = 0;
  for (const c of partes) { buf.set(c, p); p += c.byteLength; }
  return new TextDecoder().decode(buf);
}

class CorpoGrande extends Error {}

async function readJson(request, limite = 16 * 1024) {
  const bruto = await leCorpo(request, limite);
  if (bruto === null) throw new CorpoGrande();
  try {
    const v = JSON.parse(bruto);
    return v && typeof v === 'object' && !Array.isArray(v) ? v : null;
  } catch {
    return null;
  }
}

/* ============================== rotas ============================== */

async function handleActivate(request, env) {
  const body = await readJson(request);
  const licenseKey = texto(body && body.licenseKey, 64);
  const deviceId = texto(body && body.deviceId, 120);
  if (!licenseKey || !deviceId) return err('BAD_REQUEST', 'licenseKey e deviceId são obrigatórios');
  const deviceLabel = typeof body.deviceLabel === 'string' ? body.deviceLabel.trim().slice(0, 120) || null : null;
  const appVersion = typeof body.appVersion === 'string' ? body.appVersion.slice(0, 32) : null;

  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_key = ?')
    .bind(licenseKey.toUpperCase())
    .first();
  if (!license) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);

  const status = effectiveStatus(license);
  if (status === 'REVOKED' || status === 'SUSPENDED' || status === 'EXPIRED') {
    return err('LICENSE_NOT_ACTIVATABLE', `Licença está ${status}`, 403);
  }

  /* Um único comando, atômico no D1, faz as três coisas que antes eram
     separadas (ler a contagem, comparar, inserir):
     - só grava se a licença ainda tiver vaga (sem contar o próprio aparelho);
     - reativa um aparelho desativado (antes: erro 500 de chave primária,
       porque device_id é PRIMARY KEY e o código tentava INSERT de novo);
     - move o aparelho de outra licença para esta (mesmo navegador trocando
       de chave; antes também dava 500).
     Com a leitura separada da escrita, seis ativações simultâneas de uma
     licença de 1 aparelho passavam todas. */
  const agora = nowIso();
  const res = await env.DB.prepare(
    `INSERT INTO devices (device_id, license_id, device_label, first_seen_at, last_validated_at, deactivated_at)
     SELECT ?1, ?2, ?3, ?4, ?4, NULL
      WHERE (SELECT COUNT(*) FROM devices
              WHERE license_id = ?2 AND deactivated_at IS NULL AND device_id <> ?1) < ?5
     ON CONFLICT(device_id) DO UPDATE SET
       license_id = excluded.license_id,
       device_label = COALESCE(excluded.device_label, devices.device_label),
       last_validated_at = excluded.last_validated_at,
       deactivated_at = NULL`
  )
    .bind(deviceId, license.license_id, deviceLabel, agora, license.max_devices)
    .run();

  if (!res.meta || !res.meta.changes) {
    return err(
      'LICENSE_DEVICE_LIMIT_REACHED',
      'Esta licença já está ativada no número máximo de dispositivos permitido.',
      409
    );
  }

  if (license.status === 'PENDING') {
    await env.DB.prepare("UPDATE licenses SET status = 'ACTIVE', activated_at = ? WHERE license_id = ? AND status = 'PENDING'")
      .bind(agora, license.license_id)
      .run();
    license.status = 'ACTIVE';
    license.activated_at = agora;
  }

  const signed = await buildSignedLicense(env, license, { device_id: deviceId });
  await logEvent(env, { licenseId: license.license_id, deviceId, event: 'ACTIVATE', request, appVersion });
  return json(signed);
}

async function handleValidate(request, env) {
  const body = await readJson(request);
  const licenseId = texto(body && body.licenseId, 64);
  const deviceId = texto(body && body.deviceId, 120);
  if (!licenseId || !deviceId) return err('BAD_REQUEST', 'licenseId e deviceId são obrigatórios');

  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_id = ?').bind(licenseId).first();
  if (!license) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);

  const device = await env.DB.prepare('SELECT * FROM devices WHERE device_id = ? AND license_id = ?')
    .bind(deviceId, licenseId)
    .first();
  if (!device) return err('DEVICE_NOT_FOUND', 'Dispositivo não registrado nesta licença', 404);

  if (device.deactivated_at) {
    // dispositivo foi desativado (ex.: pelo fluxo de "gerenciar dispositivos") — assina um
    // payload dizendo isso, em vez de simplesmente recusar, para o app conseguir mostrar
    // uma mensagem clara e derrubar o estado PRO local de forma confiável.
    const payloadObj = {
      licenseId: license.license_id,
      deviceId,
      plan: license.plan,
      status: 'DEVICE_DEACTIVATED',
      maxDevices: license.max_devices,
      activeDevices: await countActiveDevices(env, license.license_id),
      expiresAt: license.expires_at,
      offlineToleranceDays: offlineToleranceDays(license),
      issuedAt: nowIso(),
    };
    const signed = await signLicensePayload(env, payloadObj);
    await logEvent(env, { licenseId: license.license_id, deviceId, event: 'VALIDATE', detail: 'device_deactivated', request });
    return json(signed);
  }

  await env.DB.prepare('UPDATE devices SET last_validated_at = ? WHERE device_id = ? AND license_id = ?')
    .bind(nowIso(), deviceId, licenseId)
    .run();
  const signed = await buildSignedLicense(env, license, device);
  await logEvent(env, { licenseId: license.license_id, deviceId, event: 'VALIDATE', detail: effectiveStatus(license), request });
  return json(signed);
}

async function handleListDevices(request, env) {
  /* POST com a chave no corpo é o caminho certo: na URL (GET), a chave de
     licença, que funciona como senha, ficava gravada em histórico do
     navegador e em logs de rede. O GET continua aceito por um tempo, só
     para versões antigas do app que ainda estejam em cache. */
  let bruto = null;
  if (request.method === 'POST') {
    const body = await readJson(request);
    bruto = body && body.licenseKey;
  } else {
    bruto = new URL(request.url).searchParams.get('licenseKey');
  }
  const licenseKey = texto(bruto, 64);
  if (!licenseKey) return err('BAD_REQUEST', 'licenseKey é obrigatório');

  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_key = ?').bind(licenseKey.toUpperCase()).first();
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
  const licenseKey = texto(body && body.licenseKey, 64);
  const deviceId = texto(body && body.deviceId, 120);
  if (!licenseKey || !deviceId) return err('BAD_REQUEST', 'licenseKey e deviceId são obrigatórios');

  const license = await env.DB.prepare('SELECT * FROM licenses WHERE license_key = ?')
    .bind(licenseKey.toUpperCase())
    .first();
  if (!license) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);

  /* O filtro por license_id vai no próprio UPDATE: a autorização não
     depende de uma leitura anterior que poderia ficar desatualizada. */
  const res = await env.DB.prepare('UPDATE devices SET deactivated_at = ? WHERE device_id = ? AND license_id = ?')
    .bind(nowIso(), deviceId, license.license_id)
    .run();
  if (!res.meta || !res.meta.changes) return err('DEVICE_NOT_FOUND', 'Dispositivo não encontrado nesta licença', 404);

  await logEvent(env, { licenseId: license.license_id, deviceId, event: 'DEACTIVATE', request });
  return json({ ok: true });
}

async function handleLicenseBySession(request, env) {
  const url = new URL(request.url);
  const sessionId = texto(url.searchParams.get('session_id'), 255);
  /* Só aceita o formato de sessão do Stripe: corta de saída qualquer
     tentativa de usar o parâmetro para outra coisa. */
  if (!sessionId || !/^cs_(live|test)_[A-Za-z0-9]{10,250}$/.test(sessionId)) return err('BAD_REQUEST', 'session_id inválido');

  const license = await env.DB.prepare('SELECT license_key, plan, status FROM licenses WHERE stripe_checkout_session_id = ?')
    .bind(sessionId)
    .first();
  if (!license) return err('NOT_READY', 'Ainda processando o pagamento — tente novamente em alguns segundos.', 404);

  return json(license);
}

/* Reenvio da chave para o e-mail da compra ("perdi minha chave").

   Regra de ouro deste endpoint: ele NUNCA revela se um e-mail existe ou
   não na base. A resposta é sempre a mesma, exista ou não. Quem pede o
   reenvio não recebe a chave na resposta — a chave sai só pelo e-mail
   que já estava gravado na compra. Assim o endpoint não vira nem uma
   ferramenta de descobrir clientes, nem uma de roubar chaves alheias:
   pedir o reenvio do e-mail de outra pessoa só faz o e-mail chegar a
   ela, não a quem pediu. */
async function handleResendLicense(request, env) {
  const body = await readJson(request);
  const email = body && typeof body.email === 'string' ? body.email.trim().toLowerCase() : '';

  const generic = json({
    ok: true,
    message: 'Se houver uma compra com esse e-mail, a chave foi reenviada para ele.',
  });

  if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || email.length > 254) return generic;

  const rows = await env.DB.prepare(
    `SELECT l.license_id, l.license_key, l.plan
       FROM licenses l
       JOIN customers c ON c.customer_id = l.customer_id
      WHERE lower(c.email) = ?
        AND l.status IN ('PENDING','ACTIVE')
      ORDER BY l.purchased_at DESC
      LIMIT 3`
  )
    .bind(email)
    .all();

  /* Intervalo mínimo de 15 minutos por licença. Sem ele, qualquer pessoa
     que soubesse o e-mail de um cliente podia disparar dezenas de e-mails
     para ele e esgotar a cota grátis do serviço de e-mail. */
  const limite = new Date(Date.now() - 15 * 60 * 1000).toISOString();
  const list = (rows && rows.results) || [];
  for (const lic of list) {
    const hist = await env.DB.prepare(
      `SELECT SUM(CASE WHEN created_at > ?2 THEN 1 ELSE 0 END) AS recentes, COUNT(*) AS dia
         FROM activation_log WHERE license_id = ?1 AND event = 'RESEND' AND created_at > ?3`
    )
      .bind(lic.license_id, limite, new Date(Date.now() - 24 * 3600 * 1000).toISOString())
      .first();
    /* No máximo 1 a cada 15 minutos e 4 por dia, por licença. */
    if (hist && ((hist.recentes || 0) > 0 || (hist.dia || 0) >= 4)) continue;
    const mail = await sendLicenseEmail(env, { to: email, licenseKey: lic.license_key, plan: lic.plan });
    await logEvent(env, {
      licenseId: lic.license_id,
      event: 'RESEND',
      detail: `reenvio pedido pelo cliente: ${mail.ok ? 'OK' : 'FALHOU'} — ${mail.detail}`,
      request,
    });
  }

  /* E-mail sem licença: não grava nada. Antes gravava uma linha por pedido,
     o que deixava qualquer um encher o activation_log e consumir a cota
     diária de escritas do D1 gratuito com e-mails inventados. */
  return generic;
}

/* ---- Stripe webhook ---- */

/* Janela de tolerância da assinatura, a mesma das bibliotecas oficiais do
   Stripe. Sem ela, um evento assinado capturado uma vez podia ser
   reapresentado para sempre. */
const STRIPE_TOLERANCIA_S = 300;

async function verifyStripeSignature(request, env, rawBody) {
  /* Falha FECHADO: sem segredo configurado, nada é aceito. */
  const segredo = env.STRIPE_WEBHOOK_SECRET;
  if (typeof segredo !== 'string' || segredo.length < 16) return false;

  const sigHeader = request.headers.get('Stripe-Signature') || '';
  let t = null;
  const v1s = [];
  for (const parte of sigHeader.split(',')) {
    const i = parte.indexOf('=');
    if (i < 0) continue;
    const k = parte.slice(0, i).trim(), v = parte.slice(i + 1).trim();
    if (k === 't') t = v;
    else if (k === 'v1') v1s.push(v); // pode haver mais de uma durante a troca de segredo
  }
  if (!t || !/^\d{1,12}$/.test(t) || !v1s.length) return false;
  if (Math.abs(Math.floor(Date.now() / 1000) - Number(t)) > STRIPE_TOLERANCIA_S) return false;

  const key = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(segredo),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const sigBuf = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(`${t}.${rawBody}`));
  const expectedHex = [...new Uint8Array(sigBuf)].map((b) => b.toString(16).padStart(2, '0')).join('');
  return v1s.some((v) => iguaisTempoConstante(v, expectedHex));
}

/* De qual plano é cada Payment Link.
   O painel do Stripe NÃO permite definir metadata num Payment Link — isso
   só existe pela API, que exigiria uma chave secreta em circulação. Então
   identificamos o plano pelo ID do próprio link, que vem dentro do evento
   `checkout.session.completed`.
   Isso na verdade é melhor que metadata: o ID é o link que a pessoa
   realmente usou para pagar, então não tem como ficar dessincronizado de
   uma configuração que alguém esqueceu de atualizar. */
const PAYMENT_LINK_PLAN = {
  // Full (backend: profissional) — 3 dispositivos
  plink_1U7j2HFV7byOqCPVUEBAGC4v: 'profissional', // EUR 6,29
  plink_1U7iweFV7byOqCPVaw5oBSRg: 'profissional', // USD 6,99
  plink_1TwrkmFV7byOqCPVgTy4U7DN: 'profissional', // BRL 29,90
  // Pro (backend: apoiador) — 1 dispositivo
  plink_1U7j0rFV7byOqCPVAsglGAaw: 'apoiador', // EUR 2,09
  plink_1U7iuRFV7byOqCPVICXHBGiT: 'apoiador', // USD 2,29
  plink_1Twrf7FV7byOqCPVq0ZTZRp8: 'apoiador', // BRL 9,90
};

/* Links que NÃO são venda. O link de doação usa quantidade ajustável
   sobre R$ 1,00 e o Adaptive Pricing converte para a moeda de quem paga,
   então o valor final pode cair, por coincidência, em cima de um preço
   antigo (R$ 1,00 x 12 ≈ US$ 2,29, que era o preço do Apoiador). Sem
   esta lista, uma doação viraria licença e o doador receberia uma chave
   que ele não pediu. */
const PAYMENT_LINK_DOACAO = new Set([
  'plink_1UIFpyFV7byOqCPV8WFWVlcA', // Apoio ao GIF Local (doação)
]);

/* A dedução por valor foi REMOVIDA. Ela existia para pegar um link de
   venda novo que alguém esquecesse de mapear — mas desde setembro/2026
   não se vende mais nada, e o único link novo que existe é de doação.
   Manter a dedução só criaria licenças por acidente. Agora, um checkout
   que não esteja explicitamente em PAYMENT_LINK_PLAN não emite licença
   nenhuma: apenas registra no log. */

/* ====================== entrega da chave por e-mail ======================

   Antes disto, a chave só aparecia na página obrigado.html logo depois do
   pagamento. Quem fechasse a aba antes dela carregar tinha pago e ficado
   sem nada — sem segundo canal, sem autoatendimento. Este é o segundo
   canal.

   Deliberadamente tolerante a falha: se o e-mail não sair, a licença JÁ
   está gravada e a página de obrigado continua funcionando. O erro vai
   para o activation_log e nunca sobe como exceção — se este handler
   lançasse, o Stripe re-tentaria o webhook e criaríamos uma segunda
   licença para a mesma compra.

   Sem RESEND_API_KEY configurado, tudo aqui vira no-op silencioso: o
   comportamento volta a ser exatamente o de antes, nada quebra. */

const PLAN_LABEL = { apoiador: 'Pro', profissional: 'Full', empresa: 'Full' };

function licenseEmailHtml(licenseKey, plan) {
  const label = PLAN_LABEL[plan] || 'Pro';
  return `<!doctype html><html><body style="margin:0;padding:24px;background:#f6f7f9;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1a1a1a">
<div style="max-width:520px;margin:0 auto;background:#fff;border-radius:12px;padding:32px">
<h1 style="margin:0 0 8px;font-size:20px">Sua chave do GIF Local ${label}</h1>
<p style="margin:0 0 24px;font-size:15px;line-height:1.5;color:#444">Obrigado pela compra. Guarde este e-mail: ele é o seu comprovante e a sua chave.</p>
<div style="background:#f2f4f7;border:1px solid #dfe3e8;border-radius:8px;padding:16px;text-align:center;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:19px;letter-spacing:1px;font-weight:600">${licenseKey}</div>
<p style="margin:24px 0 8px;font-size:15px;line-height:1.5;color:#444">Para ativar:</p>
<ol style="margin:0 0 24px;padding-left:20px;font-size:15px;line-height:1.7;color:#444">
<li>Abra <a href="https://giflocal.pages.dev/" style="color:#4f46e5">giflocal.pages.dev</a></li>
<li>Clique em <strong>Ativar licença</strong></li>
<li>Cole a chave acima</li>
</ol>
<p style="margin:0;font-size:13px;line-height:1.6;color:#777">O GIF Local processa tudo no seu navegador — suas imagens e vídeos nunca são enviados para a internet. A chave só libera os recursos; ela não muda isso.</p>
</div></body></html>`;
}

function licenseEmailText(licenseKey, plan) {
  const label = PLAN_LABEL[plan] || 'Pro';
  return [
    `Sua chave do GIF Local ${label}`,
    '',
    'Obrigado pela compra. Guarde este e-mail: ele e o seu comprovante e a sua chave.',
    '',
    `CHAVE: ${licenseKey}`,
    '',
    'Para ativar:',
    '1. Abra https://giflocal.pages.dev/',
    '2. Clique em "Ativar licenca"',
    '3. Cole a chave acima',
  ].join('\n');
}

/* Devolve { ok, detail } — NUNCA lança. Quem chama decide o que logar. */
async function sendLicenseEmail(env, { to, licenseKey, plan }) {
  if (!env.RESEND_API_KEY) return { ok: false, detail: 'RESEND_API_KEY ausente — envio desativado' };
  if (!to) return { ok: false, detail: 'sem e-mail de destino' };

  const from = env.LICENSE_EMAIL_FROM || 'GIF Local <onboarding@resend.dev>';
  const label = PLAN_LABEL[plan] || 'Pro';

  try {
    const res = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${env.RESEND_API_KEY}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        from,
        to: [to],
        subject: `Sua chave do GIF Local ${label}`,
        html: licenseEmailHtml(licenseKey, plan),
        text: licenseEmailText(licenseKey, plan),
      }),
    });
    if (!res.ok) {
      const body = await res.text();
      return { ok: false, detail: `Resend HTTP ${res.status}: ${body.slice(0, 300)}` };
    }
    return { ok: true, detail: 'e-mail enviado' };
  } catch (e) {
    return { ok: false, detail: `falha de rede ao chamar Resend: ${String(e && e.message ? e.message : e)}` };
  }
}

function resolvePlanFromSession(session) {
  const link = typeof session.payment_link === 'string'
    ? session.payment_link
    : session.payment_link && session.payment_link.id;

  /* Doação nunca vira licença, aconteça o que acontecer. Verificado
     antes de tudo, inclusive antes do metadata. */
  if (link && PAYMENT_LINK_DOACAO.has(link)) return null;

  const meta = session.metadata && session.metadata.plan;
  if (meta && PLAN_DEVICES[meta]) return meta;

  if (link && PAYMENT_LINK_PLAN[link]) return PAYMENT_LINK_PLAN[link];

  return null;
}

async function handleStripeWebhook(request, env) {
  const rawBody = await leCorpo(request, 256 * 1024);
  if (rawBody === null) return err('PAYLOAD_TOO_LARGE', 'Corpo da requisição grande demais.', 413);
  const validSig = await verifyStripeSignature(request, env, rawBody);
  if (!validSig) return err('INVALID_SIGNATURE', 'Assinatura do webhook inválida', 400);

  let event;
  try { event = JSON.parse(rawBody); } catch { return err('BAD_REQUEST', 'Corpo inválido', 400); }
  if (!event || !event.data || !event.data.object) return json({ ok: true });

  if (event.type === 'checkout.session.completed' || event.type === 'checkout.session.async_payment_succeeded') {
    const session = event.data.object;

    /* Só emite com o dinheiro confirmado. Com boleto (e outros meios
       assíncronos), o Stripe manda checkout.session.completed com
       payment_status = "unpaid" assim que o boleto é GERADO; o pagamento
       chega depois, em async_payment_succeeded. Antes, a licença saía na
       geração do boleto, e bastava não pagar. */
    const pago = session.payment_status === 'paid' || session.payment_status === 'no_payment_required';
    if (!pago) {
      await logEvent(env, { event: 'WEBHOOK', detail: `${event.type} aguardando pagamento — session=${session.id}`, request });
      return json({ ok: true, pending: true });
    }

    const email = session.customer_details?.email || session.customer_email;
    const plan = resolvePlanFromSession(session);
    if (!email || !plan || !PLAN_DEVICES[plan]) {
      /* Registramos o suficiente para emitir a licença à mão depois: sem
         isso, um cliente que pagou ficaria sem licença e sem rastro do
         porquê. */
      await logEvent(env, {
        event: 'WEBHOOK',
        detail: `${event.type} sem plano identificado — session=${session.id} link=${session.payment_link || '-'} valor=${session.amount_total} ${session.currency}`,
        request,
      });
      return json({ ok: true, warning: 'plano não identificado — ver activation_log e emitir manualmente' });
    }

    /* Idempotência: uma sessão de checkout gera no máximo UMA licença.
       O Stripe reenvia o webhook quando não recebe resposta a tempo, e o
       mesmo evento assinado podia ser reapresentado; cada reenvio criava
       uma licença nova e mandava mais um e-mail. */
    const jaExiste = await env.DB.prepare('SELECT license_id FROM licenses WHERE stripe_checkout_session_id = ?')
      .bind(session.id)
      .first();
    if (jaExiste) return json({ ok: true, duplicate: true });

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
    /* O NOT EXISTS no próprio INSERT fecha a janela entre a checagem acima
       e a gravação, caso duas entregas do mesmo evento cheguem juntas. */
    const ins = await env.DB.prepare(
      `INSERT INTO licenses
        (license_id, license_key, customer_id, product, plan, status, max_devices,
         stripe_checkout_session_id, stripe_payment_id, purchased_at)
       SELECT ?1, ?2, ?3, 'giflocal', ?4, 'PENDING', ?5, ?6, ?7, ?8
        WHERE NOT EXISTS (SELECT 1 FROM licenses WHERE stripe_checkout_session_id = ?6)`
    )
      .bind(licenseId, licenseKey, customer.customer_id, plan, PLAN_DEVICES[plan], session.id, session.payment_intent || null, nowIso())
      .run();
    if (!ins.meta || !ins.meta.changes) return json({ ok: true, duplicate: true });

    await logEvent(env, { licenseId, event: 'WEBHOOK', detail: `${event.type} -> licença criada`, request });

    /* Segundo canal de entrega. Falha aqui não pode derrubar o webhook:
       a licença já existe e a página de obrigado já consegue buscá-la. */
    const mail = await sendLicenseEmail(env, { to: email, licenseKey, plan });
    await logEvent(env, {
      licenseId,
      event: 'WEBHOOK',
      detail: `entrega por e-mail: ${mail.ok ? 'OK' : 'FALHOU'} — ${mail.detail}`,
      request,
    });
  } else if (event.type === 'checkout.session.async_payment_failed') {
    await logEvent(env, { event: 'WEBHOOK', detail: `pagamento assíncrono falhou — session=${event.data.object.id}`, request });
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
  /* Falha FECHADO. Antes: sem ADMIN_TOKEN configurado, a comparação era
     com a string "Bearer undefined", e quem mandasse exatamente isso virava
     administrador. Em produção o segredo existe (conferido na auditoria),
     mas a proteção não pode depender disso. */
  const token = env.ADMIN_TOKEN;
  if (typeof token !== 'string' || token.length < 24) return false;
  const auth = request.headers.get('Authorization') || '';
  return iguaisTempoConstante(auth, `Bearer ${token}`);
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
  const reason = typeof body.reason === 'string' ? body.reason.slice(0, 200) : null;
  const res = await env.DB.prepare('UPDATE licenses SET status = ?, revoked_at = ?, revoked_reason = ? WHERE license_id = ?')
    .bind(body.status, body.status === 'ACTIVE' ? null : nowIso(), reason, licenseId)
    .run();
  if (!res.meta || !res.meta.changes) return err('LICENSE_NOT_FOUND', 'Licença não encontrada', 404);
  await logEvent(env, { licenseId, event: 'REVOKE', detail: `admin -> ${body.status} (${reason || ''})`, request });
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

/* Teto grátis por IP. O limite grátis é por deviceId, e o deviceId é
   gerado pelo próprio navegador: trocando-o a cada pedido, uma só pessoa
   consumia sozinha o teto global do dia e derrubava a IA para todos.
   3 por IP deixa folga para uma casa ou escritório com IP compartilhado. */
const AI_FREE_PER_IP = 3;

/* Linhas especiais na mesma tabela ai_usage (sem mudar o esquema):
   '*global*' guarda o total do dia; 'ip:<hash>' o total grátis por IP. */
const AI_GLOBAL_SUBJECT = '*global*';

const BACKEND_PLAN_TO_TIER = { apoiador: 'pro', profissional: 'full', empresa: 'full' };

function todayKey() {
  return nowIso().slice(0, 10);
}

/* Descobre o tier de quem está pedindo. Sem licença = free. A chave de
   licença é conferida no banco (nunca confiamos num "plano" enviado pelo
   navegador, que qualquer um poderia forjar). */
async function resolveAiSubject(env, body) {
  const deviceId = texto(body && body.deviceId, 120) || '';
  let tier = 'free';
  let subject = 'dev:' + (deviceId || 'anon');
  const key = texto(body && body.licenseKey, 64);
  if (key) {
    const lic = await env.DB.prepare('SELECT license_id, plan, status FROM licenses WHERE license_key = ?')
      .bind(key.toUpperCase())
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
  const global = await env.DB.prepare('SELECT count FROM ai_usage WHERE day = ? AND subject = ?')
    .bind(day, AI_GLOBAL_SUBJECT)
    .first();
  return { day, used: (mine && mine.count) || 0, globalUsed: (global && global.count) || 0 };
}

/* Reserva uma unidade de cota de forma ATÔMICA: o incremento só acontece
   se o contador ainda estiver abaixo do limite, num único comando. Antes
   o código lia o uso, comparava e só depois incrementava; dez pedidos
   simultâneos liam todos "0 usados" e passavam todos. */
async function reservaCota(env, day, subject, limite) {
  const r = await env.DB.prepare(
    `INSERT INTO ai_usage (day, subject, count) VALUES (?1, ?2, 1)
     ON CONFLICT(day, subject) DO UPDATE SET count = count + 1 WHERE ai_usage.count < ?3`
  )
    .bind(day, subject, limite)
    .run();
  return !!(r.meta && r.meta.changes);
}

/* Na primeira geração do dia, o contador global nasce com a soma do que
   já foi usado hoje (dia da implantação: havia uso registrado só por
   pessoa). INSERT OR IGNORE: depois da primeira vez não grava nada. */
async function semeiaGlobal(env, day) {
  await env.DB.prepare(
    `INSERT OR IGNORE INTO ai_usage (day, subject, count)
     SELECT ?1, '*global*', COALESCE(SUM(count), 0) FROM ai_usage
      WHERE day = ?1 AND (subject LIKE 'dev:%' OR subject LIKE 'lic:%')`
  )
    .bind(day)
    .run();
}

async function devolveCota(env, day, subject) {
  await env.DB.prepare('UPDATE ai_usage SET count = count - 1 WHERE day = ? AND subject = ? AND count > 0')
    .bind(day, subject)
    .run();
}

async function handleAiQuota(request, env) {
  /* Aceita POST (chave no corpo) e, por compatibilidade, GET. */
  let body;
  if (request.method === 'POST') body = (await readJson(request)) || {};
  else {
    const url = new URL(request.url);
    body = { deviceId: url.searchParams.get('deviceId'), licenseKey: url.searchParams.get('licenseKey') };
  }
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
  const day = todayKey();
  const allowed = AI_DAILY_LIMIT[tier];

  /* Leitura prévia, só para recusar barato quando o dia já acabou: sem
     ela, cada pedido recusado ainda gastaria escritas no D1. A garantia de
     verdade vem das reservas atômicas logo abaixo. */
  await semeiaGlobal(env, day);
  const previa = await readAiUsage(env, subject);
  if (previa.globalUsed >= AI_GLOBAL_DAILY_CAP) {
    return json(
      { error: 'AI_SERVICE_BUSY', message: 'O limite diário de gerações com IA do serviço foi atingido. Tente de novo amanhã.' },
      429
    );
  }
  if (previa.used >= allowed) {
    return json({ error: 'AI_LIMIT_REACHED', message: 'Você usou suas gerações com IA de hoje.', tier, used: previa.used, allowed }, 429);
  }

  /* Ordem das reservas: por IP (só no grátis), por pessoa e, por último, o
     teto global. Os limites mais estreitos vêm primeiro: quem já estourou o
     próprio limite é recusado sem gravar nada, e ninguém consegue gerar
     escritas em série (reserva + devolução) para esgotar a cota diária de
     escritas do D1 gratuito. */
  const devolver = [];
  const desfaz = async () => { for (const sub of devolver) await devolveCota(env, day, sub); };
  if (tier === 'free') {
    const h = await hashIp(env, request.headers.get('CF-Connecting-IP') || '');
    if (h) {
      const ipSubject = 'ip:' + h;
      if (!(await reservaCota(env, day, ipSubject, AI_FREE_PER_IP))) {
        return json({ error: 'AI_LIMIT_REACHED', message: 'Você usou suas gerações com IA de hoje.', tier, used: allowed, allowed }, 429);
      }
      devolver.push(ipSubject);
    }
  }
  if (!(await reservaCota(env, day, subject, allowed))) {
    await desfaz();
    return json({ error: 'AI_LIMIT_REACHED', message: 'Você usou suas gerações com IA de hoje.', tier, used: allowed, allowed }, 429);
  }
  devolver.push(subject);
  if (!(await reservaCota(env, day, AI_GLOBAL_SUBJECT, AI_GLOBAL_DAILY_CAP))) {
    await desfaz();
    return json(
      { error: 'AI_SERVICE_BUSY', message: 'O limite diário de gerações com IA do serviço foi atingido. Tente de novo amanhã.' },
      429
    );
  }
  const { used } = await readAiUsage(env, subject);

  /* A cota é gasta ANTES de gerar: uma geração que falhar custa um crédito,
     o que é preferível a um limite que não segura. */
  let result;
  try {
    result = await env.AI.run(AI_MODEL, { prompt, steps: AI_STEPS });
  } catch (e) {
    console.error('AI.run falhou:', e && e.message ? e.message : e);
    return err('AI_FAILED', 'Não foi possível gerar a imagem agora. Tente de novo em instantes.', 502);
  }

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
    console.error('AI.run: formato inesperado', typeof result);
    return err('AI_UNEXPECTED_OUTPUT', 'O modelo respondeu num formato inesperado.', 502);
  }

  return json({
    ok: true,
    image: imageB64,
    model: AI_MODEL,
    tier,
    used,
    allowed,
    remaining: Math.max(0, allowed - used),
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
      return withCors(new Response(null, { headers: corsHeaders() }), request);
    }

    try {
      return withCors(await route(request, env, url), request);
    } catch (e) {
      if (e instanceof CorpoGrande) return withCors(err('PAYLOAD_TOO_LARGE', 'Corpo da requisição grande demais.', 413), request);
      /* O detalhe vai para o log do Worker (visível só no painel); quem
         chamou recebe uma mensagem genérica. Antes a mensagem interna do
         JavaScript ou do D1 voltava inteira na resposta. */
      console.error('erro interno', url.pathname, e && e.stack ? e.stack : e);
      return withCors(err('INTERNAL_ERROR', 'Erro interno. Tente novamente em instantes.', 500), request);
    }
  },
};

/* Corpo máximo por rota. No plano gratuito o Worker tem 10 ms de CPU por
   pedido; um JSON de dezenas de MB derrubaria a requisição no parse. Um
   evento do Stripe tem poucos KB. */
const LIMITE_CORPO = { '/webhooks/stripe': 256 * 1024 };
const LIMITE_CORPO_PADRAO = 16 * 1024;

async function route(request, env, url) {
  if (request.method === 'POST') {
    const tam = Number(request.headers.get('Content-Length') || '0');
    if (tam > (LIMITE_CORPO[url.pathname] || LIMITE_CORPO_PADRAO)) return err('PAYLOAD_TOO_LARGE', 'Corpo da requisição grande demais.', 413);
  }

  if (url.pathname === '/activate' && request.method === 'POST') return await handleActivate(request, env);
  if (url.pathname === '/validate' && request.method === 'POST') return await handleValidate(request, env);
  if (url.pathname === '/license/devices' && (request.method === 'POST' || request.method === 'GET')) return await handleListDevices(request, env);
  if (url.pathname === '/device/deactivate' && request.method === 'POST') return await handleDeactivateDevice(request, env);
  if (url.pathname === '/license-by-session' && request.method === 'GET') return await handleLicenseBySession(request, env);
  if (url.pathname === '/license/resend' && request.method === 'POST') return await handleResendLicense(request, env);
  if (url.pathname === '/webhooks/stripe' && request.method === 'POST') return await handleStripeWebhook(request, env);
  if (url.pathname === '/ai/generate' && request.method === 'POST') return await handleAiGenerate(request, env);
  if (url.pathname === '/ai/quota' && (request.method === 'POST' || request.method === 'GET')) return await handleAiQuota(request, env);

  const adminMatch = url.pathname.match(/^\/admin\/license\/([^/]+)(\/status)?$/);
  if (adminMatch && request.method === 'GET' && !adminMatch[2]) return await handleAdminGetLicense(request, env, adminMatch[1]);
  if (adminMatch && request.method === 'POST' && adminMatch[2] === '/status')
    return await handleAdminSetStatus(request, env, adminMatch[1]);

  return err('NOT_FOUND', 'Rota não encontrada', 404);
}

