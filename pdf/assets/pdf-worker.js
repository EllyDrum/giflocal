/* GIF Local, ferramentas de PDF: trabalho pesado fora da tela.
   Este worker roda no navegador do usuário. Ele não faz nenhuma chamada
   de rede: recebe os bytes do arquivo, devolve os bytes do resultado.
   Biblioteca: @cantoo/pdf-lib (MIT), fork do pdf-lib com suporte a PDF
   criptografado (abre com a senha correta). */
'use strict';
importScripts('/pdf/vendor/pdf-lib-2.11.1/pdf-lib.min.js');

const { PDFDocument, PDFName, PDFRawStream, PDFArray, PDFNumber, PDFDict, PDFRef, EncryptedPDFError, decodePDFRawStream, degrees, rgb, StandardFonts } = self.PDFLib;

const PRODUTOR = 'GIF Local (giflocal.com)';

function progresso(id, feito, total, etapa) {
  self.postMessage({ id, tipo: 'progresso', feito, total, etapa });
}

class Falha extends Error {
  constructor(codigo, detalhe) { super(codigo); this.codigo = codigo; this.detalhe = detalhe || null; }
}

/* Traduz erros das bibliotecas para códigos que a interface sabe explicar. */
function classifica(e) {
  if (e instanceof Falha) return e;
  const msg = String((e && e.message) || e || '');
  if (e instanceof RangeError || /allocation failed|out of memory|Invalid array length|Array buffer/i.test(msg)) return new Falha('MEMORIA');
  if (e instanceof EncryptedPDFError || /encrypted/i.test(msg)) return new Falha('CRIPTOGRAFADO');
  if (/Password incorrect|NEEDS PASSWORD/i.test(msg)) return new Falha('SENHA_ERRADA');
  if (/not supported|unsupported|Unknown (security|encryption)|filter/i.test(msg) && /crypt|security|encrypt/i.test(msg)) return new Falha('PROTECAO_INCOMPATIVEL');
  if (/Failed to parse|No PDF header|Expected|Invalid|parse/i.test(msg)) return new Falha('CORROMPIDO');
  return new Falha('INESPERADO', msg.slice(0, 200));
}

/* Qualquer falha ao ler a estrutura do arquivo (ex.: PDF cortado no meio
   do download) vira CORROMPIDO. Criptografia, senha e memória mantêm o
   próprio diagnóstico. */
async function abre(bytes, opcoes) {
  try {
    const doc = await PDFDocument.load(bytes, Object.assign({ updateMetadata: false, throwOnInvalidObject: false }, opcoes || {}));
    doc.getPageCount();
    return doc;
  } catch (e) {
    const c = classifica(e);
    if (['CRIPTOGRAFADO', 'SENHA_ERRADA', 'MEMORIA', 'PROTECAO_INCOMPATIVEL'].includes(c.codigo)) throw e;
    throw new Falha('CORROMPIDO');
  }
}

/* Estado de proteção de um PDF:
   - aberto: sem criptografia;
   - restrito: abre sem senha, mas o autor definiu restrições de uso
     (só a senha de permissões, do proprietário, autoriza removê-las);
   - senha: precisa da senha de abertura. */
async function estadoProtecao(bytes) {
  try {
    const doc = await abre(bytes);
    return { estado: 'aberto', paginas: doc.getPageCount() };
  } catch (e) {
    if (!(e instanceof EncryptedPDFError) && !/encrypted/i.test(String(e && e.message))) throw classifica(e);
  }
  try {
    const doc = await abre(bytes, { password: '' });
    return { estado: 'restrito', paginas: doc.getPageCount() };
  } catch (e) {
    const c = classifica(e);
    if (c.codigo === 'SENHA_ERRADA' || c.codigo === 'CRIPTOGRAFADO') return { estado: 'senha', paginas: null };
    throw c;
  }
}

async function carregaParaEditar(bytes, indice) {
  try {
    return await abre(bytes);
  } catch (e) {
    const c = classifica(e);
    if (c.codigo === 'CRIPTOGRAFADO') throw new Falha('CRIPTOGRAFADO', { indice });
    throw new Falha(c.codigo, Object.assign({ indice }, c.detalhe ? { msg: c.detalhe } : {}));
  }
}

function marca(doc) {
  doc.setProducer(PRODUTOR);
  doc.setCreator(PRODUTOR);
}

/* ---------------- operações ---------------- */

async function opInfo(p) {
  return estadoProtecao(p.bytes);
}

async function opDesbloquear(p, id) {
  progresso(id, 0, 3, 'verificando');
  const info = await estadoProtecao(p.bytes);
  if (info.estado === 'aberto') throw new Falha('SEM_PROTECAO');
  const senha = typeof p.senha === 'string' ? p.senha : '';
  if (info.estado === 'restrito' && senha === '') throw new Falha('PRECISA_SENHA_PROPRIETARIO');
  if (senha === '') throw new Falha('PRECISA_SENHA');
  let doc;
  progresso(id, 1, 3, 'abrindo');
  try {
    doc = await abre(p.bytes, { password: senha });
  } catch (e) {
    const c = classifica(e);
    throw c.codigo === 'CRIPTOGRAFADO' ? new Falha('SENHA_ERRADA') : c;
  }
  progresso(id, 2, 3, 'gravando');
  const saida = await doc.save({ useObjectStreams: true });
  return { bytes: saida, paginas: doc.getPageCount() };
}

async function opUnir(p, id) {
  const destino = await PDFDocument.create();
  marca(destino);
  const total = p.arquivos.length;
  for (let i = 0; i < total; i++) {
    progresso(id, i, total, 'unindo');
    const origem = await carregaParaEditar(p.arquivos[i], i);
    const paginas = await destino.copyPages(origem, origem.getPageIndices());
    for (const pg of paginas) destino.addPage(pg);
    p.arquivos[i] = null; // libera a memória do arquivo já copiado
  }
  progresso(id, total, total, 'gravando');
  const bytes = await destino.save({ useObjectStreams: true });
  return { bytes, paginas: destino.getPageCount() };
}

async function opDividir(p, id) {
  const origem = await carregaParaEditar(p.bytes, 0);
  const n = origem.getPageCount();
  const saidas = [];
  for (let g = 0; g < p.grupos.length; g++) {
    progresso(id, g, p.grupos.length, 'separando');
    const indices = p.grupos[g].filter((i) => Number.isInteger(i) && i >= 0 && i < n);
    if (!indices.length) continue;
    const novo = await PDFDocument.create();
    marca(novo);
    const paginas = await novo.copyPages(origem, indices);
    for (const pg of paginas) novo.addPage(pg);
    saidas.push({ bytes: await novo.save({ useObjectStreams: true }), paginas: indices.length, rotulo: p.rotulos ? p.rotulos[g] : null });
  }
  return { saidas, paginasOrigem: n };
}

const TAMANHOS = { a4: [595.28, 841.89], carta: [612, 792] };

async function opImagensParaPdf(p, id) {
  const doc = await PDFDocument.create();
  marca(doc);
  const { tamanho, orientacao, margem, ajuste } = p.opcoes || {};
  /* "Preencher" ocupa a página inteira (sangria total): margem não se aplica. */
  const m = ajuste === 'preencher' && tamanho !== 'imagem' ? 0 : margem === 'pequena' ? 18 : margem === 'grande' ? 42 : 0;
  for (let i = 0; i < p.imagens.length; i++) {
    progresso(id, i, p.imagens.length, 'montando');
    const img = p.imagens[i];
    const emb = img.tipo === 'png' ? await doc.embedPng(img.bytes) : await doc.embedJpg(img.bytes);
    /* tamanho natural: pixels a 96 ppp viram pontos (1/72 pol.) */
    const natW = (img.largura || emb.width) * 0.75, natH = (img.altura || emb.height) * 0.75;
    let pw, ph;
    if (img.pagina) {
      /* página com tamanho exato (compressão máxima: a imagem é a própria página) */
      const pagina = doc.addPage(img.pagina);
      pagina.drawImage(emb, { x: 0, y: 0, width: img.pagina[0], height: img.pagina[1] });
      p.imagens[i] = null;
      continue;
    }
    if (tamanho === 'imagem') { pw = natW + 2 * m; ph = natH + 2 * m; }
    else {
      [pw, ph] = TAMANHOS[tamanho] || TAMANHOS.a4;
      const paisagem = orientacao === 'paisagem' || (orientacao === 'auto' && natW > natH);
      if (paisagem) [pw, ph] = [ph, pw];
    }
    const pagina = doc.addPage([pw, ph]);
    const aw = pw - 2 * m, ah = ph - 2 * m;
    let w, h;
    if (tamanho === 'imagem') { w = natW; h = natH; }
    else {
      const esc = ajuste === 'preencher' ? Math.max(aw / natW, ah / natH) : Math.min(aw / natW, ah / natH);
      w = natW * esc; h = natH * esc;
    }
    const x = m + (aw - w) / 2, y = m + (ah - h) / 2;
    /* Em "preencher", o que passa da página fica fora da área visível. */
    pagina.drawImage(emb, { x, y, width: w, height: h });
    p.imagens[i] = null;
  }
  progresso(id, p.imagens.length, p.imagens.length, 'gravando');
  return { bytes: await doc.save({ useObjectStreams: true }), paginas: doc.getPageCount() };
}

/* ---------------- compressão ---------------- */

const NIVEIS = {
  baixa: { imagens: false },
  media: { imagens: true, ladoMax: 2000, qualidade: 0.78, flate: false },
  alta: { imagens: true, ladoMax: 1400, qualidade: 0.62, flate: true },
};

function nome(v) { return v instanceof PDFName ? v.asString() : null; }

function resolve(ctx, v) { return v instanceof PDFRef ? ctx.lookup(v) : v; }

/* Só mexe em imagens que sabemos recodificar sem mudar a cor: 8 bits,
   RGB ou cinza (inclusive ICCBased de 1 ou 3 canais). CMYK, paletas,
   máscaras e matrizes /Decode ficam como estão. */
function espacoDeCor(ctx, dict) {
  const cs = resolve(ctx, dict.get(PDFName.of('ColorSpace')));
  const n = nome(cs);
  if (n === '/DeviceRGB') return 3;
  if (n === '/DeviceGray') return 1;
  if (cs instanceof PDFArray && nome(cs.get(0)) === '/ICCBased') {
    const perfil = resolve(ctx, cs.get(1));
    const dictPerfil = perfil && perfil.dict;
    const N = dictPerfil && dictPerfil.get(PDFName.of('N'));
    if (N instanceof PDFNumber && (N.asNumber() === 3 || N.asNumber() === 1)) return N.asNumber();
  }
  return 0;
}

function filtros(dict) {
  const f = dict.get(PDFName.of('Filter'));
  if (!f) return [];
  if (f instanceof PDFName) return [f.asString()];
  if (f instanceof PDFArray) return f.asArray().map(nome);
  return ['?'];
}

/* ASCII85 (usado por alguns geradores, como o ReportLab, antes do DCT). */
function decodificaA85(bytes) {
  let z = 0;
  for (let k = 0; k < bytes.length; k++) if (bytes[k] === 0x7A) z++;
  const saida = new Uint8Array(Math.ceil(bytes.length / 5) * 4 + z * 4 + 8);
  let o = 0, tupla = 0, n = 0, i = 0;
  if (bytes[0] === 0x3C && bytes[1] === 0x7E) i = 2; // "<~"
  for (; i < bytes.length; i++) {
    const c = bytes[i];
    if (c === 0x7E) break; // "~>"
    if (c <= 0x20) continue;
    if (c === 0x7A && n === 0) { o += 4; continue; } // "z" = quatro zeros
    if (c < 0x21 || c > 0x75) throw new Error('A85 inválido');
    tupla = tupla * 85 + (c - 33);
    if (++n === 5) {
      saida[o++] = (tupla >>> 24) & 255; saida[o++] = (tupla >>> 16) & 255; saida[o++] = (tupla >>> 8) & 255; saida[o++] = tupla & 255;
      tupla = 0; n = 0;
    }
  }
  if (n > 0) {
    for (let k = n; k < 5; k++) tupla = tupla * 85 + 84;
    const b = [(tupla >>> 24) & 255, (tupla >>> 16) & 255, (tupla >>> 8) & 255, tupla & 255];
    for (let k = 0; k < n - 1; k++) saida[o++] = b[k];
  }
  return saida.subarray(0, o);
}

async function recodifica(bitmap, ladoMax, qualidade) {
  let w = bitmap.width, h = bitmap.height;
  const esc = Math.min(1, ladoMax / Math.max(w, h));
  w = Math.max(1, Math.round(w * esc)); h = Math.max(1, Math.round(h * esc));
  const tela = new OffscreenCanvas(w, h);
  const c = tela.getContext('2d');
  c.fillStyle = '#fff';
  c.fillRect(0, 0, w, h);
  c.imageSmoothingQuality = 'high';
  c.drawImage(bitmap, 0, 0, w, h);
  bitmap.close && bitmap.close();
  const blob = await tela.convertToBlob({ type: 'image/jpeg', quality: qualidade });
  return { bytes: new Uint8Array(await blob.arrayBuffer()), w, h };
}

async function opComprimir(p, id) {
  const cfg = NIVEIS[p.nivel] || NIVEIS.media;
  const doc = await carregaParaEditar(p.bytes, 0);
  const ctx = doc.context;
  const stats = { imagens: 0, recodificadas: 0, semSuporte: 0 };
  const podeImagens = cfg.imagens && typeof OffscreenCanvas !== 'undefined' && typeof createImageBitmap === 'function';
  if (cfg.imagens && !podeImagens) stats.semSuporteNavegador = true;

  if (podeImagens) {
    const alvos = [];
    for (const [ref, obj] of ctx.enumerateIndirectObjects()) {
      if (!(obj instanceof PDFRawStream)) continue;
      const d = obj.dict;
      if (nome(d.get(PDFName.of('Subtype'))) !== '/Image') continue;
      stats.imagens++;
      alvos.push([ref, obj]);
    }
    for (let i = 0; i < alvos.length; i++) {
      progresso(id, i, alvos.length, 'imagens');
      const [ref, obj] = alvos[i];
      const d = obj.dict;
      try {
        const bpc = d.get(PDFName.of('BitsPerComponent'));
        const canais = espacoDeCor(ctx, d);
        const fs = filtros(d);
        const w = d.get(PDFName.of('Width')), h = d.get(PDFName.of('Height'));
        const W = w instanceof PDFNumber ? w.asNumber() : 0, H = h instanceof PDFNumber ? h.asNumber() : 0;
        if (!canais || !(bpc instanceof PDFNumber) || bpc.asNumber() !== 8 || d.get(PDFName.of('Decode')) || d.get(PDFName.of('ImageMask')) || W * H < 160000) { stats.semSuporte++; continue; }
        let bitmap = null;
        const ehJpeg = fs[fs.length - 1] === '/DCTDecode' && fs.slice(0, -1).every((f) => f === '/ASCII85Decode');
        if (ehJpeg && fs.length <= 2) {
          const jpeg = fs.length === 2 ? decodificaA85(obj.contents) : obj.contents;
          bitmap = await createImageBitmap(new Blob([jpeg], { type: 'image/jpeg' }));
        } else if (cfg.flate && fs.length === 1 && fs[0] === '/FlateDecode' && !d.get(PDFName.of('DecodeParms'))) {
          const px = decodePDFRawStream(obj).decode();
          if (px.length < W * H * canais) { stats.semSuporte++; continue; }
          const rgba = new Uint8ClampedArray(W * H * 4);
          for (let k = 0, s = 0; k < W * H; k++) {
            if (canais === 3) { rgba[k * 4] = px[s]; rgba[k * 4 + 1] = px[s + 1]; rgba[k * 4 + 2] = px[s + 2]; s += 3; }
            else { rgba[k * 4] = rgba[k * 4 + 1] = rgba[k * 4 + 2] = px[s]; s += 1; }
            rgba[k * 4 + 3] = 255;
          }
          bitmap = await createImageBitmap(new ImageData(rgba, W, H));
        } else { stats.semSuporte++; continue; }
        const r = await recodifica(bitmap, cfg.ladoMax, cfg.qualidade);
        if (r.bytes.length >= obj.contents.length * 0.9) continue; // não compensa
        const novo = ctx.stream(r.bytes, {
          Type: 'XObject', Subtype: 'Image', Width: r.w, Height: r.h,
          ColorSpace: 'DeviceRGB', BitsPerComponent: 8, Filter: 'DCTDecode',
        });
        const smask = d.get(PDFName.of('SMask'));
        if (smask) novo.dict.set(PDFName.of('SMask'), smask);
        const interp = d.get(PDFName.of('Interpolate'));
        if (interp) novo.dict.set(PDFName.of('Interpolate'), interp);
        ctx.assign(ref, novo);
        stats.recodificadas++;
      } catch (e) {
        stats.semSuporte++;
      }
    }
  }
  progresso(id, 1, 1, 'gravando');
  const bytes = await doc.save({ useObjectStreams: true });
  return { bytes, stats, paginas: doc.getPageCount() };
}


/* ---------------- organizar / girar ---------------- */
/* p.paginas: [{ indice, rotacao }] na ordem final (páginas ausentes são excluídas) */
async function opOrganizar(p, id) {
  const origem = await carregaComSenha(p.bytes, p.senha);
  const destino = await PDFDocument.create();
  marca(destino);
  const total = p.paginas.length;
  if (!total) throw new Falha('SEM_PAGINAS');
  const copiadas = await destino.copyPages(origem, p.paginas.map((x) => x.indice));
  copiadas.forEach((pg, i) => {
    const giro = ((p.paginas[i].rotacao || 0) % 360 + 360) % 360;
    if (giro) pg.setRotation(degrees((pg.getRotation().angle + giro) % 360));
    destino.addPage(pg);
    if (i % 20 === 0) progresso(id, i, total, 'organizando');
  });
  progresso(id, total, total, 'gravando');
  const bytes = await destino.save({ useObjectStreams: true });
  return { bytes, paginas: destino.getPageCount() };
}
/* p.giro: 90|180|270; p.indices: null (todas) ou lista */
async function opGirar(p, id) {
  const doc = await carregaComSenha(p.bytes, p.senha);
  marca(doc);
  const pgs = doc.getPages();
  const alvo = p.indices ? new Set(p.indices) : null;
  pgs.forEach((pg, i) => { if (!alvo || alvo.has(i)) pg.setRotation(degrees((pg.getRotation().angle + p.giro) % 360)); });
  progresso(id, 1, 1, 'gravando');
  const bytes = await doc.save({ useObjectStreams: true });
  return { bytes, paginas: pgs.length };
}
async function carregaComSenha(bytes, senha) {
  try { return await abre(bytes, senha ? { password: senha } : undefined); }
  catch (e) { const c = classifica(e); if (c.codigo === 'CRIPTOGRAFADO') throw new Falha(senha ? 'SENHA_ERRADA' : 'PRECISA_SENHA'); throw c; }
}

/* ---------------- proteger (AES-256) ---------------- */
async function opProteger(p, id) {
  const doc = await carregaComSenha(p.bytes, p.senhaAtual);
  marca(doc);
  const perm = p.permissoes || {};
  progresso(id, 1, 2, 'criptografando');
  doc.encrypt({
    userPassword: p.senhaAbertura || '',
    ownerPassword: p.senhaProprietario || p.senhaAbertura || '',
    permissions: {
      printing: perm.imprimir === false ? false : 'highResolution',
      modifying: perm.editar !== false,
      copying: perm.copiar !== false,
      annotating: perm.editar !== false,
      fillingForms: true,
      contentAccessibility: true,
      documentAssembly: perm.editar !== false,
    },
  });
  progresso(id, 2, 2, 'gravando');
  const bytes = await doc.save({ useObjectStreams: true });
  return { bytes, paginas: doc.getPageCount() };
}

/* ---------------- remover metadados ---------------- */
async function opMetadados(p, id) {
  const doc = await carregaComSenha(p.bytes, p.senha);
  const antes = { titulo: seguro(() => doc.getTitle()), autor: seguro(() => doc.getAuthor()), assunto: seguro(() => doc.getSubject()), palavras: seguro(() => doc.getKeywords()), criador: seguro(() => doc.getCreator()), produtor: seguro(() => doc.getProducer()), criacao: seguro(() => doc.getCreationDate()), modificacao: seguro(() => doc.getModificationDate()) };
  const achados = Object.keys(antes).filter((k) => antes[k]);
  /* 1. dicionário Info */
  const info = doc.context.trailerInfo.Info;
  if (info instanceof PDFRef) doc.context.delete(info);
  delete doc.context.trailerInfo.Info;
  /* 2. XMP no catálogo e nas páginas */
  let xmp = 0;
  const limpaDict = (d) => { if (d && d.has && d.has(PDFName.of('Metadata'))) { d.delete(PDFName.of('Metadata')); xmp++; } };
  limpaDict(doc.catalog);
  for (const pg of doc.getPages()) limpaDict(pg.node);
  /* 3. identificadores do trailer */
  delete doc.context.trailerInfo.ID;
  progresso(id, 1, 1, 'gravando');
  const bytes = await doc.save({ useObjectStreams: true, updateMetadata: false });
  return { bytes, paginas: doc.getPageCount(), removidos: achados, xmp };
}
function seguro(fn) { try { return fn() || null; } catch (e) { return null; } }

/* ---------------- marca-d'água e numeração ---------------- */
const SUBST = { '–': '-', '—': '-', '‘': "'", '’': "'", '“': '"', '”': '"', '…': '...', ' ': ' ' };
function textoPadrao(fonte, t) {
  const ok = new Set(fonte.getCharacterSet());
  let out = '', trocados = 0;
  for (const ch of String(t)) {
    const cp = ch.codePointAt(0);
    if (ok.has(cp)) { out += ch; continue; }
    if (SUBST[ch]) { out += SUBST[ch]; continue; }
    const base = ch.normalize('NFD')[0];
    if (base && ok.has(base.codePointAt(0))) { out += base; continue; }
    trocados++; out += '?';
  }
  return { texto: out, trocados };
}
const corHex = (h) => { const m = /^#?([0-9a-f]{6})$/i.exec(h || ''); if (!m) return rgb(0.5, 0.5, 0.5); const n = parseInt(m[1], 16); return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); };

/* p: { texto, tamanho, opacidade, cor, posicao: 'diagonal'|'centro'|'rodape'|'mosaico', indices? } */
async function opMarcaDagua(p, id) {
  const doc = await carregaComSenha(p.bytes, p.senha);
  marca(doc);
  const fonte = await doc.embedFont(StandardFonts.HelveticaBold);
  const { texto, trocados } = textoPadrao(fonte, p.texto || '');
  if (!texto.trim()) throw new Falha('TEXTO_VAZIO');
  const pgs = doc.getPages();
  const alvo = p.indices ? new Set(p.indices) : null;
  const cor = corHex(p.cor);
  const opacity = Math.max(0.05, Math.min(1, +p.opacidade || 0.3));
  pgs.forEach((pg, i) => {
    if (alvo && !alvo.has(i)) return;
    const { width: W, height: H } = pg.getSize();
    const rot = pg.getRotation().angle;
    /* tamanho relativo à página: p.tamanho em % da largura */
    let size = Math.max(8, (W * (p.tamanho || 40)) / 100 / Math.max(1, fonte.widthOfTextAtSize(texto, 100) / 100));
    const tw = fonte.widthOfTextAtSize(texto, size);
    const desenha = (x, y, ang) => pg.drawText(texto, { x, y, size, font: fonte, color: cor, opacity, rotate: degrees(ang + rot) });
    if (p.posicao === 'mosaico') {
      size = Math.max(8, size * 0.45);
      const tw2 = fonte.widthOfTextAtSize(texto, size);
      const passoX = tw2 + size * 2, passoY = size * 5;
      for (let y = -H * 0.2; y < H * 1.2; y += passoY) for (let x = -W * 0.2 - (y / passoY % 2) * passoX / 2; x < W * 1.2; x += passoX) pg.drawText(texto, { x, y, size, font: fonte, color: cor, opacity, rotate: degrees(30 + rot) });
    } else if (p.posicao === 'rodape') desenha((W - tw) / 2, size * 1.2, 0);
    else if (p.posicao === 'topo') desenha((W - tw) / 2, H - size * 2, 0);
    else if (p.posicao === 'centro') desenha((W - tw) / 2, (H - size) / 2, 0);
    else { const ang = Math.atan2(H, W) * 180 / Math.PI; const cx = W / 2, cy = H / 2; const rad = ang * Math.PI / 180; desenha(cx - (tw / 2) * Math.cos(rad) + (size / 2) * Math.sin(rad), cy - (tw / 2) * Math.sin(rad) - (size / 2) * Math.cos(rad), ang); }
    if (i % 25 === 0) progresso(id, i, pgs.length, 'marcando');
  });
  progresso(id, pgs.length, pgs.length, 'gravando');
  const bytes = await doc.save({ useObjectStreams: true });
  return { bytes, paginas: pgs.length, trocados };
}

/* p: { formato: '{n}'|'{n}/{total}'|'Página {n}'..., posicao: 'rodape-centro'|'rodape-direita'|'rodape-esquerda'|'topo-centro'|'topo-direita'|'topo-esquerda', inicio, tamanho, margem, primeira (pular capa?) } */
async function opNumerar(p, id) {
  const doc = await carregaComSenha(p.bytes, p.senha);
  marca(doc);
  const fonte = await doc.embedFont(StandardFonts.Helvetica);
  const pgs = doc.getPages();
  const total = pgs.length;
  const inicio = parseInt(p.inicio, 10) || 1;
  const size = Math.max(6, Math.min(36, +p.tamanho || 11));
  const margem = Math.max(6, Math.min(80, +p.margem || 28));
  const cor = corHex(p.cor || '#333333');
  const modelo = String(p.formato || '{n}');
  let trocados = 0;
  pgs.forEach((pg, i) => {
    if (p.pularPrimeira && i === 0) return;
    const n = inicio + i - (p.pularPrimeira ? 1 : 0);
    const bruto = modelo.replace(/\{n\}/g, String(n)).replace(/\{total\}/g, String(p.pularPrimeira ? total - 1 : total));
    const t = textoPadrao(fonte, bruto); trocados += t.trocados;
    const { width: W, height: H } = pg.getSize();
    const tw = fonte.widthOfTextAtSize(t.texto, size);
    const [vert, horiz] = (p.posicao || 'rodape-centro').split('-');
    const x = horiz === 'esquerda' ? margem : horiz === 'direita' ? W - margem - tw : (W - tw) / 2;
    const y = vert === 'topo' ? H - margem - size * 0.7 : margem;
    pg.drawText(t.texto, { x, y, size, font: fonte, color: cor });
    if (i % 25 === 0) progresso(id, i, total, 'numerando');
  });
  progresso(id, total, total, 'gravando');
  const bytes = await doc.save({ useObjectStreams: true });
  return { bytes, paginas: total, trocados };
}

const OPERACOES = {
  organizar: opOrganizar,
  girar: opGirar,
  proteger: opProteger,
  metadados: opMetadados,
  marcaDagua: opMarcaDagua,
  numerar: opNumerar,
  info: opInfo,
  desbloquear: opDesbloquear,
  unir: opUnir,
  dividir: opDividir,
  imagensParaPdf: opImagensParaPdf,
  comprimir: opComprimir,
};

self.onmessage = async (ev) => {
  const { id, op, dados } = ev.data || {};
  const fn = OPERACOES[op];
  if (!fn) { self.postMessage({ id, tipo: 'erro', codigo: 'INESPERADO' }); return; }
  try {
    const r = await fn(dados, id);
    const transferir = [];
    if (r && r.bytes && r.bytes.buffer) transferir.push(r.bytes.buffer);
    if (r && r.saidas) for (const s of r.saidas) transferir.push(s.bytes.buffer);
    self.postMessage({ id, tipo: 'ok', resultado: r }, transferir);
  } catch (e) {
    const c = classifica(e);
    self.postMessage({ id, tipo: 'erro', codigo: c.codigo, detalhe: c.detalhe });
  }
};
