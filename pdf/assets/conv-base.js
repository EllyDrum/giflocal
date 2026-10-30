/* GIF Local, conversões de documentos: peças comuns.
   ZIP (leitura e escrita), XML e utilidades. Tudo roda no navegador:
   arquivos do Office (.docx, .xlsx, .pptx) são pacotes ZIP com XML dentro,
   e são lidos e gerados aqui, sem bibliotecas externas e sem rede. */

export class ErroConversao extends Error {
  constructor(codigo, detalhe) { super(codigo); this.codigo = codigo; this.detalhe = detalhe || null; }
}

/* ---------------- CRC32 ---------------- */
const TABELA = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
export function crc32(b) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < b.length; i++) c = TABELA[(c ^ b[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

const enc = new TextEncoder();
export const utf8 = (s) => enc.encode(s);

async function transforma(bytes, stream, limite) {
  const leitor = new Blob([bytes]).stream().pipeThrough(stream).getReader();
  const partes = [];
  let total = 0;
  for (;;) {
    const { done, value } = await leitor.read();
    if (done) break;
    total += value.length;
    /* proteção contra "bombas de ZIP": um arquivo pequeno que se expande
       para gigabytes travaria o navegador */
    if (limite && total > limite) { try { await leitor.cancel(); } catch (e) { /* ignora */ } throw new ErroConversao('ARQUIVO_INVALIDO'); }
    partes.push(value);
  }
  const out = new Uint8Array(total);
  let p = 0;
  for (const v of partes) { out.set(v, p); p += v.length; }
  return out;
}

/* ---------------- ZIP: escrita (com compressão deflate) ---------------- */
/* itens: [{ nome, dados: Uint8Array|string, comprimir?: boolean }] */
export async function criaZip(itens, tipo) {
  const partes = [], central = [];
  let desloc = 0;
  const d = new Date();
  const hora = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
  const data = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  for (const it of itens) {
    const nb = utf8(it.nome);
    const bruto = typeof it.dados === 'string' ? utf8(it.dados) : it.dados;
    const crc = crc32(bruto);
    let corpo = bruto, metodo = 0;
    if (it.comprimir !== false && bruto.length > 64 && typeof CompressionStream !== 'undefined') {
      const z = await transforma(bruto, new CompressionStream('deflate-raw'));
      if (z.length < bruto.length) { corpo = z; metodo = 8; }
    }
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, metodo, true);
    lh.setUint16(10, hora, true); lh.setUint16(12, data, true); lh.setUint32(14, crc, true);
    lh.setUint32(18, corpo.length, true); lh.setUint32(22, bruto.length, true); lh.setUint16(26, nb.length, true);
    partes.push(new Uint8Array(lh.buffer), nb, corpo);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, metodo, true); ch.setUint16(12, hora, true); ch.setUint16(14, data, true); ch.setUint32(16, crc, true);
    ch.setUint32(20, corpo.length, true); ch.setUint32(24, bruto.length, true); ch.setUint16(28, nb.length, true);
    ch.setUint32(42, desloc, true);
    central.push(new Uint8Array(ch.buffer), nb);
    desloc += 30 + nb.length + corpo.length;
  }
  let tamCentral = 0;
  for (const c of central) tamCentral += c.length;
  const fim = new DataView(new ArrayBuffer(22));
  fim.setUint32(0, 0x06054b50, true); fim.setUint16(8, itens.length, true); fim.setUint16(10, itens.length, true);
  fim.setUint32(12, tamCentral, true); fim.setUint32(16, desloc, true);
  return new Blob([...partes, ...central, new Uint8Array(fim.buffer)], { type: tipo || 'application/zip' });
}

/* ---------------- ZIP: leitura ---------------- */
const LIMITE_ENTRADA = 300 * 1048576; // 300 MB descompactados por parte
const LIMITE_TOTAL = 600 * 1048576;

export async function abreZip(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new ErroConversao('ARQUIVO_INVALIDO');
  const n = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const entradas = new Map();
  const dec = new TextDecoder();
  for (let k = 0; k < n; k++) {
    if (p + 46 > bytes.length || dv.getUint32(p, true) !== 0x02014b50) throw new ErroConversao('ARQUIVO_INVALIDO');
    const metodo = dv.getUint16(p + 10, true);
    const comp = dv.getUint32(p + 20, true);
    const tam = dv.getUint32(p + 24, true);
    const ln = dv.getUint16(p + 28, true), le = dv.getUint16(p + 30, true), lc = dv.getUint16(p + 32, true);
    const loc = dv.getUint32(p + 42, true);
    const nome = dec.decode(bytes.subarray(p + 46, p + 46 + ln));
    entradas.set(nome, { metodo, comp, tam, loc });
    p += 46 + ln + le + lc;
  }
  let consumido = 0;
  const cache = new Map();
  async function le(nome) {
    if (cache.has(nome)) return cache.get(nome);
    const e = entradas.get(nome);
    if (!e) return null;
    if (e.loc + 30 > bytes.length || dv.getUint32(e.loc, true) !== 0x04034b50) throw new ErroConversao('ARQUIVO_INVALIDO');
    const ini = e.loc + 30 + dv.getUint16(e.loc + 26, true) + dv.getUint16(e.loc + 28, true);
    const corpo = bytes.subarray(ini, ini + e.comp);
    let out;
    if (e.metodo === 0) out = corpo;
    else if (e.metodo === 8) out = await transforma(corpo, new DecompressionStream('deflate-raw'), Math.min(LIMITE_ENTRADA, LIMITE_TOTAL - consumido));
    else throw new ErroConversao('ARQUIVO_INVALIDO');
    consumido += out.length;
    if (consumido > LIMITE_TOTAL) throw new ErroConversao('ARQUIVO_INVALIDO');
    cache.set(nome, out);
    return out;
  }
  return {
    nomes: () => [...entradas.keys()],
    tem: (nome) => entradas.has(nome),
    bytes: le,
    async texto(nome) { const b = await le(nome); return b ? new TextDecoder('utf-8').decode(b) : null; },
    async xml(nome) {
      const s = await this.texto(nome);
      if (s == null) return null;
      const d = new DOMParser().parseFromString(s, 'application/xml');
      if (d.getElementsByTagName('parsererror').length) throw new ErroConversao('ARQUIVO_INVALIDO');
      return d;
    },
  };
}

/* ---------------- XML ---------------- */
export function escXml(s) {
  return String(s)
    // caracteres de controle não são permitidos em XML 1.0
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
/* Navegação por nome local (ignora o prefixo do namespace: w:, a:, p:...). */
export function filhos(el, nome) {
  const out = [];
  if (!el) return out;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (!nome || c.localName === nome) out.push(c);
  return out;
}
export function filho(el, nome) {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === nome) return c;
  return null;
}
export function caminho(el, ...nomes) {
  let a = el;
  for (const n of nomes) { a = filho(a, n); if (!a) return null; }
  return a;
}
export function descendentes(el, nome) {
  const out = [];
  if (!el) return out;
  const todos = el.getElementsByTagName('*');
  for (let i = 0; i < todos.length; i++) if (todos[i].localName === nome) out.push(todos[i]);
  return out;
}
export function atr(el, nome) {
  if (!el || !el.attributes) return null;
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (a.localName === nome) return a.value;
  }
  return null;
}
/* Caminho relativo de uma relação (ex.: "../media/image1.png" a partir de "word/") */
export function resolveCaminho(base, alvo) {
  if (!alvo) return null;
  if (alvo.startsWith('/')) return alvo.slice(1);
  const partes = base.split('/').filter(Boolean);
  for (const s of alvo.split('/')) {
    if (s === '..') partes.pop();
    else if (s !== '.' && s) partes.push(s);
  }
  return partes.join('/');
}
/* Lê um arquivo .rels e devolve id -> { alvo, tipo, externo } */
export async function leRelacoes(zip, arquivo) {
  const pasta = arquivo.includes('/') ? arquivo.slice(0, arquivo.lastIndexOf('/') + 1) : '';
  const nomeRels = pasta + '_rels/' + arquivo.slice(pasta.length) + '.rels';
  const mapa = new Map();
  const d = await zip.xml(nomeRels);
  if (!d) return mapa;
  for (const r of descendentes(d.documentElement, 'Relationship')) {
    const externo = r.getAttribute('TargetMode') === 'External';
    mapa.set(r.getAttribute('Id'), { alvo: externo ? r.getAttribute('Target') : resolveCaminho(pasta, r.getAttribute('Target')), tipo: (r.getAttribute('Type') || '').split('/').pop(), externo });
  }
  return mapa;
}

/* ---------------- imagens ---------------- */
export function tipoImagemBytes(b) {
  if (!b || b.length < 12) return null;
  if (b[0] === 0xFF && b[1] === 0xD8) return 'jpg';
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
  if (b[0] === 0x42 && b[1] === 0x4D) return 'bmp';
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57) return 'webp';
  if ((b[0] === 0x49 && b[1] === 0x49) || (b[0] === 0x4D && b[1] === 0x4D)) return 'tiff';
  return null;
}
/* GIF, BMP e WebP não entram direto num PDF: o navegador decodifica e
   regrava como PNG. EMF/WMF (desenhos do Office) não são suportados. */
export async function imagemParaPdf(bytes) {
  const tipo = tipoImagemBytes(bytes);
  if (tipo === 'jpg' || tipo === 'png') return { bytes, tipo };
  if (tipo === 'gif' || tipo === 'bmp' || tipo === 'webp') {
    try {
      const bmp = await createImageBitmap(new Blob([bytes]));
      const c = document.createElement('canvas');
      c.width = bmp.width; c.height = bmp.height;
      c.getContext('2d').drawImage(bmp, 0, 0);
      bmp.close && bmp.close();
      const blob = await new Promise((ok) => c.toBlob(ok, 'image/png'));
      c.width = 0; c.height = 0;
      if (!blob) return null;
      return { bytes: new Uint8Array(await blob.arrayBuffer()), tipo: 'png' };
    } catch (e) { return null; }
  }
  return null;
}

export function pausa() { return new Promise((r) => setTimeout(r, 0)); }
