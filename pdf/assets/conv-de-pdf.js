/* GIF Local, conversões a partir de PDF: Word, Excel, PowerPoint e texto.
   O PDF é lido pelo PDF.js no navegador. O texto vem com posição e
   tamanho de cada trecho; daqui saem linhas, parágrafos e tabelas, que
   viram arquivos do Office montados aqui mesmo (sem servidor). */
import { criaZip, escXml, pausa } from './conv-base.js?v=4';
import { TEMA, MESTRE } from './conv-pptx-modelo.js?v=4';

const NEGRITO = /bold|black|heavy|semibold|demibold|extrabold|ultrabold|,bd|-bd\b/i;
const ITALICO = /italic|oblique|,it\b|-it\b/i;
const FONTES = [
  [/arial|helvetica|liberationsans|arimo|nimbussans/i, 'Arial'], [/times|liberationserif|tinos|nimbusroman/i, 'Times New Roman'],
  [/calibri|carlito/i, 'Calibri'], [/cambria|caladea/i, 'Cambria'], [/courier|liberationmono|cousine|nimbusmono/i, 'Courier New'],
  [/verdana/i, 'Verdana'], [/georgia/i, 'Georgia'], [/tahoma/i, 'Tahoma'], [/segoe/i, 'Segoe UI'], [/garamond/i, 'Garamond'],
  [/trebuchet/i, 'Trebuchet MS'], [/roboto/i, 'Roboto'], [/opensans/i, 'Open Sans'], [/dejavusans/i, 'DejaVu Sans'], [/aptos/i, 'Aptos'],
];
function familiaDe(bruto) {
  const s = String(bruto || '').replace(/^[A-Z]{6}\+/, '').replace(/\s+/g, '');
  for (const [re, n] of FONTES) if (re.test(s)) return n;
  return null;
}

/* ================================================================
   Extração: trechos de texto (e imagens) de uma página
   ================================================================ */
export async function extraiPagina(pdfjs, page, op = {}) {
  const vp = page.getViewport({ scale: 1 });
  const W = vp.width, H = vp.height;
  let ops = null;
  if (op.fontes || op.imagens) { try { ops = await page.getOperatorList(); } catch (e) { ops = null; } }
  const tc = await page.getTextContent();
  const cacheFonte = new Map();
  function fonte(nome) {
    if (cacheFonte.has(nome)) return cacheFonte.get(nome);
    let real = '', f = null;
    try { if (ops && page.commonObjs.has(nome)) { f = page.commonObjs.get(nome); real = (f && f.name) || ''; } } catch (e) { /* sem fonte */ }
    const st = (tc.styles && tc.styles[nome]) || {};
    const fam = String(st.fontFamily || '');
    const r = {
      negrito: NEGRITO.test(real) || !!(f && f.bold && !/light/i.test(real)),
      italico: ITALICO.test(real) || !!(f && f.italic),
      familia: familiaDe(real) || (fam === 'monospace' ? 'Courier New' : fam === 'serif' ? 'Times New Roman' : null),
    };
    cacheFonte.set(nome, r);
    return r;
  }
  const itens = [], girados = [];
  for (const it of tc.items) {
    if (typeof it.str !== 'string') continue;
    /* marcadores de lista em fontes de símbolos (Symbol, Wingdings) chegam como caracteres privados */
    const s = it.str.replace(/[\u0000-\u0008\u000B-\u001F]/g, '').replace(/[\uF0B7\uF0A7\uF076\uF0D8\uF0FC\uF06C\uF06E\uF0A8\uF0B2\uF0E8]/g, '•');
    if (!s.trim()) continue;
    const m = pdfjs.Util.transform(vp.transform, it.transform);
    const tam = Math.hypot(m[2], m[3]);
    if (!(tam >= 1) || tam > 400) continue;
    const ang = Math.atan2(m[1], m[0]);
    const f = fonte(it.fontName);
    const reg = { s, x0: m[4], x1: m[4] + Math.max(0, it.width) * Math.cos(ang), y: m[5], tam, negrito: f.negrito, italico: f.italico, familia: f.familia };
    if (Math.abs(ang) > 0.05) { girados.push(reg); continue; }
    if (reg.x1 < reg.x0) [reg.x0, reg.x1] = [reg.x1, reg.x0];
    itens.push(reg);
  }
  const imagens = [];
  if (op.imagens && ops) {
    const O = pdfjs.OPS;
    let ctm = [1, 0, 0, 1, 0, 0];
    const pilha = [];
    for (let i = 0; i < ops.fnArray.length; i++) {
      const fn = ops.fnArray[i], a = ops.argsArray[i];
      if (fn === O.save) pilha.push(ctm);
      else if (fn === O.restore) ctm = pilha.pop() || [1, 0, 0, 1, 0, 0];
      else if (fn === O.transform) ctm = pdfjs.Util.transform(ctm, a);
      else if (fn === O.paintFormXObjectBegin) { pilha.push(ctm); if (a && Array.isArray(a[0]) && a[0].length === 6) ctm = pdfjs.Util.transform(ctm, a[0]); }
      else if (fn === O.paintFormXObjectEnd) ctm = pilha.pop() || [1, 0, 0, 1, 0, 0];
      else if (fn === O.paintImageXObject || fn === O.paintInlineImageXObject) {
        const m = pdfjs.Util.transform(vp.transform, ctm);
        const xs = [m[4], m[0] + m[4], m[2] + m[4], m[0] + m[2] + m[4]];
        const ys = [m[5], m[1] + m[5], m[3] + m[5], m[1] + m[3] + m[5]];
        const caixa = { x0: Math.max(0, Math.min(...xs)), x1: Math.min(W, Math.max(...xs)), y0: Math.max(0, Math.min(...ys)), y1: Math.min(H, Math.max(...ys)) };
        if (caixa.x1 - caixa.x0 < 12 || caixa.y1 - caixa.y0 < 12) continue;
        imagens.push({ caixa, ref: a[0], inline: fn === O.paintInlineImageXObject });
      }
    }
  }
  const letras = itens.reduce((n, it) => n + it.s.trim().length, 0) + girados.reduce((n, it) => n + it.s.trim().length, 0);
  const grandeImagem = imagens.some((im) => (im.caixa.x1 - im.caixa.x0) * (im.caixa.y1 - im.caixa.y0) > 0.5 * W * H);
  return { W, H, itens, girados, imagens, letras, escaneada: letras < 8 && (grandeImagem || imagens.length > 0) };
}

/* Pixels de uma imagem do PDF (já decodificada pelo PDF.js) num canvas. */
export async function canvasDaImagem(page, im) {
  let obj = null;
  try {
    if (im.inline) obj = im.ref;
    else {
      const repo = String(im.ref).startsWith('g_') ? page.commonObjs : page.objs;
      if (repo.has(im.ref)) obj = repo.get(im.ref);
    }
  } catch (e) { obj = null; }
  if (!obj) return null;
  const w = obj.width, h = obj.height;
  if (!w || !h || w * h > 40e6) return null;
  const c = document.createElement('canvas');
  c.width = w; c.height = h;
  const ctx = c.getContext('2d');
  if (obj.bitmap) { ctx.drawImage(obj.bitmap, 0, 0); return c; }
  if (!obj.data) return null;
  const id = ctx.createImageData(w, h);
  const d = id.data, s = obj.data;
  if (obj.kind === 3) d.set(s.subarray(0, d.length));
  else if (obj.kind === 2) { for (let i = 0, j = 0; i < d.length; i += 4, j += 3) { d[i] = s[j]; d[i + 1] = s[j + 1]; d[i + 2] = s[j + 2]; d[i + 3] = 255; } }
  else if (obj.kind === 1) {
    const bpl = (w + 7) >> 3;
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const bit = (s[y * bpl + (x >> 3)] >> (7 - (x & 7))) & 1;
      const v = bit ? 255 : 0, i = (y * w + x) * 4;
      d[i] = d[i + 1] = d[i + 2] = v; d[i + 3] = 255;
    }
  } else return null;
  ctx.putImageData(id, 0, 0);
  return c;
}
export async function canvasParaBytes(c, maxLado) {
  let fonte = c;
  if (maxLado && Math.max(c.width, c.height) > maxLado) {
    const k = maxLado / Math.max(c.width, c.height);
    fonte = document.createElement('canvas');
    fonte.width = Math.max(1, Math.round(c.width * k)); fonte.height = Math.max(1, Math.round(c.height * k));
    fonte.getContext('2d').drawImage(c, 0, 0, fonte.width, fonte.height);
  }
  /* transparência pede PNG; foto sem transparência fica menor em JPG */
  let alfa = false;
  try {
    const amostra = fonte.getContext('2d').getImageData(0, 0, fonte.width, fonte.height).data;
    const passo = Math.max(4, Math.floor(amostra.length / 4 / 20000) * 4);
    for (let i = 3; i < amostra.length; i += passo) if (amostra[i] < 250) { alfa = true; break; }
  } catch (e) { alfa = true; }
  const png = alfa || fonte.width * fonte.height < 40000;
  const blob = await new Promise((ok) => fonte.toBlob(ok, png ? 'image/png' : 'image/jpeg', 0.85));
  const largura = fonte.width, altura = fonte.height;
  if (fonte !== c) { fonte.width = 0; fonte.height = 0; }
  if (!blob) return null;
  return { bytes: new Uint8Array(await blob.arrayBuffer()), ext: png ? 'png' : 'jpeg', largura, altura };
}

/* ================================================================
   Linhas e segmentos
   ================================================================ */
function juntaRuns(runs) {
  const out = [];
  for (const r of runs) {
    const u = out[out.length - 1];
    if (u && u.negrito === r.negrito && u.italico === r.italico && Math.abs(u.tam - r.tam) < 0.6 && u.familia === r.familia) u.texto += r.texto;
    else out.push(Object.assign({}, r));
  }
  return out;
}
const MARCADOR_SO = /^([•●▪◦‣∙·■□➢►✓o\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z][.)]|[ivxIVX]{1,4}[.)])$/;
export function montaLinhas(itens) {
  const ord = itens.slice().sort((a, b) => a.y - b.y || a.x0 - b.x0);
  const linhas = [];
  for (const it of ord) {
    let alvo = null;
    for (let k = linhas.length - 1; k >= Math.max(0, linhas.length - 3); k--) {
      const L = linhas[k];
      if (Math.abs(L.y - it.y) <= 0.42 * Math.min(L.tam, it.tam)) { alvo = L; break; }
    }
    if (!alvo) { alvo = { y: it.y, tam: it.tam, itens: [] }; linhas.push(alvo); }
    alvo.itens.push(it);
    if (it.tam > alvo.tam) alvo.tam = it.tam;
  }
  for (const L of linhas) {
    L.itens.sort((a, b) => a.x0 - b.x0);
    const segs = [];
    let seg = null, ant = null;
    for (const it of L.itens) {
      if (ant && Math.abs(it.x0 - ant.x0) < 0.6 && it.s === ant.s) continue; // texto repetido (negrito simulado)
      const gap = ant ? it.x0 - ant.x1 : 0;
      const larguraLetra = ant ? (ant.x1 - ant.x0) / Math.max(1, ant.s.length) : L.tam * 0.5;
      if (!seg || gap > Math.max(L.tam * 0.95, 5)) {
        seg = { x0: it.x0, x1: it.x1, runs: [] };
        segs.push(seg);
      } else if (gap > Math.min(larguraLetra * 0.28, L.tam * 0.14)) {
        const ult = seg.runs[seg.runs.length - 1];
        if (ult && !/\s$/.test(ult.texto) && !/^\s/.test(it.s)) ult.texto += ' ';
      }
      seg.runs.push({ texto: it.s, negrito: it.negrito, italico: it.italico, tam: it.tam, familia: it.familia });
      seg.x1 = Math.max(seg.x1, it.x1);
      ant = it;
    }
    for (const s of segs) {
      s.runs = juntaRuns(s.runs);
      s.runs[0].texto = s.runs[0].texto.replace(/^\s+/, '');
      const u = s.runs[s.runs.length - 1];
      u.texto = u.texto.replace(/\s+$/, '');
      s.runs = s.runs.filter((r) => r.texto);
      s.texto = s.runs.map((r) => r.texto).join('');
    }
    L.segs = segs.filter((s) => s.texto);
    /* marcador de lista separado do texto ("•   item", "1.   item") é parte da mesma frase, não uma coluna */
    if (L.segs.length >= 2 && MARCADOR_SO.test(L.segs[0].texto) && L.segs[1].x0 - L.segs[0].x1 < L.tam * 4) {
      const [m, t] = L.segs;
      L.segs.splice(0, 2, { x0: m.x0, x1: t.x1, runs: m.runs.concat([Object.assign({}, t.runs[0], { texto: ' ' })], t.runs), texto: m.texto + ' ' + t.texto, marcador: true });
    }
    L.x0 = L.segs.length ? L.segs[0].x0 : 0;
    L.x1 = L.segs.length ? L.segs[L.segs.length - 1].x1 : 0;
    L.texto = L.segs.map((s) => s.texto).join(' ');
    delete L.itens;
  }
  return linhas.filter((L) => L.segs.length);
}
/* Linhas de texto girado (raras): viram linhas simples no fim da página. */
export function linhasGiradas(girados) {
  return girados.map((g) => ({ y: 1e9, tam: g.tam, x0: 0, x1: 0, texto: g.s.trim(), segs: [{ x0: 0, x1: 0, texto: g.s.trim(), runs: [{ texto: g.s.trim(), negrito: g.negrito, italico: g.italico, tam: g.tam, familia: g.familia }] }] }));
}

/* ---------------- colunas (tabelas) ---------------- */
export function faixasDeColunas(linhas, larguraPagina) {
  const ints = [];
  for (const L of linhas) if (L.segs.length >= 2) for (const s of L.segs) if (s.x1 - s.x0 < larguraPagina * 0.45) ints.push([s.x0, s.x1]);
  ints.sort((a, b) => a[0] - b[0]);
  const f = [];
  for (const [a, b] of ints) {
    const u = f[f.length - 1];
    if (u && a <= u[1] + 1.5) u[1] = Math.max(u[1], b);
    else f.push([a, b]);
  }
  return f;
}
export function faixaDe(seg, faixas) {
  let melhor = 0, mv = -Infinity;
  for (let i = 0; i < faixas.length; i++) {
    const [a, b] = faixas[i];
    const sobre = Math.min(b, seg.x1) - Math.max(a, seg.x0);
    const v = sobre > 0 ? sobre : -Math.min(Math.abs(seg.x0 - a), Math.abs(seg.x0 - b)) - 1000;
    if (v > mv) { mv = v; melhor = i; }
  }
  return melhor;
}
function blocoAlinhado(bloco, fx) {
  let ok = 0;
  for (const L of bloco) {
    const usados = new Set(L.segs.map((s) => faixaDe(s, fx)));
    if (usados.size === L.segs.length && usados.size >= 2) ok++;
  }
  return ok / bloco.length >= 0.8 && (bloco.length >= 3 || fx.length >= 3);
}
function regioesDeTabela(linhas, W) {
  const regs = [];
  let i = 0;
  while (i < linhas.length) {
    if (linhas[i].segs.length < 2) { i++; continue; }
    let j = i;
    while (j + 1 < linhas.length && linhas[j + 1].segs.length >= 2 && linhas[j + 1].y - linhas[j].y < 3.2 * linhas[j].tam) j++;
    const bloco = linhas.slice(i, j + 1);
    if (bloco.length >= 2) {
      const fx = faixasDeColunas(bloco, W);
      if (fx.length >= 2 && fx.length <= 25 && blocoAlinhado(bloco, fx)) regs.push({ ini: i, fim: j, faixas: fx });
    }
    i = j + 1;
  }
  return regs;
}

/* ================================================================
   Números (PDF -> Excel)
   ================================================================ */
/* Decide se o documento usa vírgula decimal (1.234,56) ou ponto (1,234.56). */
export function separadorDecimal(textos) {
  let virg = 0, ponto = 0;
  for (const t of textos) {
    if (/\d\.\d{3},\d{1,4}\b/.test(t) || /(^|[^\d.,])\d+,\d{1,2}(%|\b)(?![.,]\d)/.test(t)) virg++;
    if (/\d,\d{3}\.\d{1,4}\b/.test(t) || /(^|[^\d.,])\d+\.\d{1,2}(%|\b)(?![.,]\d)/.test(t)) ponto++;
  }
  return virg >= ponto ? ',' : '.';
}
export function interpretaNumero(bruto, dec) {
  let s = String(bruto).trim();
  if (!s || s.length > 40) return null;
  let moeda = null;
  const mm = s.match(/^(R\$|US\$|\$|€|£)\s?/);
  if (mm) { moeda = mm[1]; s = s.slice(mm[0].length); }
  const mm2 = !moeda && s.match(/\s?(€|R\$)$/);
  if (mm2) { moeda = mm2[1]; s = s.slice(0, -mm2[0].length); }
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
  if (/^[-−–]/.test(s)) { neg = !neg; s = s.slice(1).trim(); }
  if (/-$/.test(s)) { neg = !neg; s = s.slice(0, -1).trim(); }
  if (!moeda) { const m3 = s.match(/^(R\$|US\$|\$|€)\s?/); if (m3) { moeda = m3[1]; s = s.slice(m3[0].length); } }
  let pct = false;
  if (/\s?%$/.test(s)) { pct = true; s = s.replace(/\s?%$/, ''); }
  const mil = dec === ',' ? '.' : ',';
  const reMil = dec === ',' ? /^\d{1,3}(\.\d{3})+(,\d+)?$/ : /^\d{1,3}(,\d{3})+(\.\d+)?$/;
  const reSimples = dec === ',' ? /^\d+(,\d+)?$/ : /^\d+(\.\d+)?$/;
  if (!reMil.test(s) && !reSimples.test(s)) return null;
  /* zeros à esquerda (códigos, CEP, matrícula) continuam texto */
  if (/^0\d/.test(s) && !moeda && !pct) return null;
  const inteiroLongo = s.replace(/\D/g, '').length > 15;
  if (inteiroLongo) return null;
  let v = parseFloat(s.split(mil).join('').replace(dec, '.'));
  if (!isFinite(v)) return null;
  if (neg) v = -v;
  const casas = (s.split(dec)[1] || '').length;
  let fmt;
  if (pct) { v /= 100; fmt = !casas ? 'pct0' : casas === 2 ? 'pct2' : 'pctN:' + Math.min(casas, 6); }
  else if (moeda) fmt = 'moeda:' + moeda;
  else if (casas) fmt = casas === 2 ? 'dec2' : 'decN:' + Math.min(casas, 6);
  else fmt = s.includes(mil) ? 'mil0' : 'int';
  return { v, fmt };
}
export function interpretaData(s, dec) {
  if (dec !== ',') return null; // dd/mm/aaaa só quando o documento é de padrão brasileiro/europeu
  const m = String(s).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const d = +m[1], mes = +m[2], a = +m[3];
  if (d < 1 || d > 31 || mes < 1 || mes > 12 || a < 1900 || a > 2200) return null;
  const ms = Date.UTC(a, mes - 1, d);
  const dt = new Date(ms);
  if (dt.getUTCDate() !== d) return null;
  return { v: ms / 86400000 + 25569, fmt: 'data' }; // serial de data do Excel
}

/* ================================================================
   PDF -> Excel
   ================================================================ */
export function paginaParaPlanilha(pag, dec) {
  const linhas = montaLinhas(pag.itens).concat(linhasGiradas(pag.girados));
  const fx = faixasDeColunas(linhas, pag.W);
  const ncol = Math.max(1, fx.length);
  const grade = [];
  let anterior = null;
  for (const L of linhas) {
    if (anterior && L.y - anterior.y > 2.6 * Math.max(anterior.tam, L.tam) && L.y < 1e8) grade.push([]);
    const linha = new Array(ncol).fill(null);
    for (const s of L.segs) {
      const c = fx.length ? faixaDe(s, fx) : 0;
      const negrito = s.runs.every((r) => r.negrito);
      if (linha[c]) linha[c].texto += ' ' + s.texto;
      else linha[c] = { texto: s.texto, negrito };
    }
    grade.push(linha.map((cel) => {
      if (!cel) return null;
      const n = interpretaNumero(cel.texto, dec) || interpretaData(cel.texto, dec);
      return n ? { v: n.v, fmt: n.fmt, negrito: cel.negrito, texto: cel.texto } : { s: cel.texto, negrito: cel.negrito };
    }));
    anterior = L;
  }
  while (grade.length && grade[grade.length - 1].every((c) => !c)) grade.pop();
  return grade;
}

const COLUNA = (n) => { let s = ''; n++; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
export { COLUNA };

function nomeDePlanilha(nome, usados) {
  let base = String(nome).replace(/[\[\]:*?/\\]/g, ' ').slice(0, 31).trim() || 'Planilha';
  let n = base, k = 2;
  while (usados.has(n.toLowerCase())) n = base.slice(0, 27) + ' (' + (k++) + ')';
  usados.add(n.toLowerCase());
  return n;
}

/* planilhas: [{ nome, grade: [[celula|null]] }] */
export async function geraXlsx(planilhas) {
  const formatosMoeda = [];
  const estilos = { int: 1, mil0: 3, dec2: 4, pct0: 9, pct2: 10, data: 14 };
  const xfs = [['0', '0']]; // [numFmtId, fontId]
  const mapaXf = new Map([['0|0', 0]]);
  const numFmts = [];
  function idFormato(fmt) {
    if (fmt in estilos) return estilos[fmt];
    if (fmt.startsWith('pctN:')) {
      const cod = '0.' + '0'.repeat(+fmt.slice(5)) + '%';
      let f = numFmts.find((x) => x.cod === cod); if (!f) { f = { id: 164 + numFmts.length, cod }; numFmts.push(f); } return f.id;
    }
    if (fmt.startsWith('decN:')) {
      const n = +fmt.slice(5); const cod = '#,##0.' + '0'.repeat(n);
      let f = numFmts.find((x) => x.cod === cod); if (!f) { f = { id: 164 + numFmts.length, cod }; numFmts.push(f); } return f.id;
    }
    if (fmt.startsWith('moeda:')) {
      const sim = fmt.slice(6); const cod = '"' + sim + '" #,##0.00;-"' + sim + '" #,##0.00';
      let f = numFmts.find((x) => x.cod === cod); if (!f) { f = { id: 164 + numFmts.length, cod }; numFmts.push(f); formatosMoeda.push(sim); } return f.id;
    }
    return 0;
  }
  function xf(fmt, negrito) {
    const id = fmt ? idFormato(fmt) : 0, fonte = negrito ? 1 : 0;
    const k = id + '|' + fonte;
    if (!mapaXf.has(k)) { mapaXf.set(k, xfs.length); xfs.push([String(id), String(fonte)]); }
    return mapaXf.get(k);
  }
  const usados = new Set();
  const folhas = [];
  for (const p of planilhas) {
    const nome = nomeDePlanilha(p.nome, usados);
    const larg = [];
    let xml = '';
    p.grade.forEach((linha, r) => {
      if (!linha.length || linha.every((c) => !c)) return;
      let cel = '';
      linha.forEach((c, k) => {
        if (!c) return;
        const ref = COLUNA(k) + (r + 1);
        const comp = c.s != null ? c.s.length : String(c.texto || c.v).length;
        larg[k] = Math.max(larg[k] || 6, Math.min(60, comp + 2));
        if (c.s != null) {
          const s = xf(null, c.negrito);
          cel += `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${escXml(c.s)}</t></is></c>`;
        } else {
          cel += `<c r="${ref}" s="${xf(c.fmt, c.negrito)}"><v>${c.v}</v></c>`;
        }
      });
      xml += `<row r="${r + 1}">${cel}</row>`;
    });
    const cols = larg.length ? '<cols>' + larg.map((w, k) => w ? `<col min="${k + 1}" max="${k + 1}" width="${w}" customWidth="1"/>` : '').join('') + '</cols>' : '';
    folhas.push({ nome, xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"${folhas.length ? '' : ' tabSelected="1"'}/></sheetViews><sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${xml}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>` });
    await pausa();
  }
  const estilosXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((f) => `<numFmt numFmtId="${f.id}" formatCode="${escXml(f.cod)}"/>`).join('')}</numFmts>` : ''}<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.map(([n, f]) => `<xf numFmtId="${n}" fontId="${f}" fillId="0" borderId="0" xfId="0"${n !== '0' ? ' applyNumberFormat="1"' : ''}${f !== '0' ? ' applyFont="1"' : ''}/>`).join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  const itens = [
    { nome: '[Content_Types].xml', dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${folhas.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>` },
    { nome: '_rels/.rels', dados: RELS_RAIZ('xl/workbook.xml') },
    { nome: 'docProps/core.xml', dados: CORE() },
    { nome: 'docProps/app.xml', dados: APP('Microsoft Excel') },
    { nome: 'xl/workbook.xml', dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${folhas.map((f, i) => `<sheet name="${escXml(f.nome)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>` },
    { nome: 'xl/_rels/workbook.xml.rels', dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${folhas.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${folhas.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
    { nome: 'xl/styles.xml', dados: estilosXml },
    ...folhas.map((f, i) => ({ nome: `xl/worksheets/sheet${i + 1}.xml`, dados: f.xml })),
  ];
  return criaZip(itens, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

const RELS_RAIZ = (alvo) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${alvo}"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
const CORE = () => {
  const agora = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>GIF Local</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${agora}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${agora}</dcterms:modified></cp:coreProperties>`;
};
const APP = (app) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>${app}</Application></Properties>`;

/* ================================================================
   PDF -> Word
   ================================================================ */
const MARCADOR = /^([•●▪◦‣∙·■□➢►✓\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z][.)]|[ivxIVX]{1,4}[.)])\s/;

function percentil(v, p) { if (!v.length) return 0; const s = v.slice().sort((a, b) => a - b); return s[Math.min(s.length - 1, Math.floor(p * s.length))]; }
function tamanhoCorpo(linhas) {
  const cont = new Map();
  for (const L of linhas) for (const s of L.segs) for (const r of s.runs) { const k = Math.round(r.tam * 2) / 2; cont.set(k, (cont.get(k) || 0) + r.texto.length); }
  let melhor = 11, mv = -1;
  for (const [k, v] of cont) if (v > mv) { mv = v; melhor = k; }
  return melhor;
}
function runsDaLinha(L) {
  const runs = [];
  L.segs.forEach((s, i) => {
    if (i) runs.push(Object.assign({}, s.runs[0], { texto: ' ' }));
    for (const r of s.runs) runs.push(Object.assign({}, r));
  });
  return runs;
}

/* Converte as linhas de uma página em blocos (parágrafos, tabelas e imagens). */
export function paginaParaBlocos(pag, figuras) {
  const linhas = montaLinhas(pag.itens);
  const corpo = tamanhoCorpo(linhas) || 11;
  const esq = percentil(linhas.map((L) => L.x0), 0.1);
  const dir = percentil(linhas.map((L) => L.x1), 0.9);
  const largura = Math.max(50, dir - esq);
  const regs = regioesDeTabela(linhas, pag.W);
  const blocos = [];
  const naTabela = new Map();
  for (const r of regs) for (let i = r.ini; i <= r.fim; i++) naTabela.set(i, r);
  let p = null, ant = null;
  const fecha = () => { if (p) { blocos.push(p); p = null; } };
  for (let i = 0; i < linhas.length; i++) {
    const L = linhas[i];
    const reg = naTabela.get(i);
    if (reg) {
      fecha(); ant = null;
      if (i === reg.ini) {
        const fx = reg.faixas;
        const grade = [];
        for (let k = reg.ini; k <= reg.fim; k++) {
          const linha = fx.map(() => null);
          for (const s of linhas[k].segs) { const c = faixaDe(s, fx); linha[c] = linha[c] ? linha[c].concat([Object.assign({}, s.runs[0], { texto: ' ' })], s.runs) : s.runs.slice(); }
          grade.push(linha);
        }
        const bordas = fx.map(([a], k) => (k ? (fx[k - 1][1] + a) / 2 : Math.min(a, esq)));
        bordas.push(Math.max(fx[fx.length - 1][1], dir));
        blocos.push({ tipo: 'tabela', y: linhas[reg.ini].y - linhas[reg.ini].tam, grade, larguras: fx.map((_, k) => Math.max(20, bordas[k + 1] - bordas[k])) });
      }
      continue;
    }
    const runs = runsDaLinha(L);
    const centro = (L.x0 + L.x1) / 2;
    const centrada = Math.abs(centro - pag.W / 2) < pag.W * 0.04 && (L.x1 - L.x0) < largura * 0.8 && L.x0 > esq + 15;
    let novo = !p;
    if (p && ant) {
      const gap = L.y - ant.y;
      if (gap > 1.55 * Math.max(ant.tam, L.tam)) novo = true;
      else if (Math.abs(L.tam - ant.tam) > 0.18 * ant.tam) novo = true;
      else if (MARCADOR.test(L.texto)) novo = true;
      else if (centrada !== p.centrada) novo = true;
      else if (ant.x1 < dir - largura * 0.18 && /[.:;!?)]$/.test(ant.texto) && !p.centrada) novo = true;
      else if (L.x0 > esq + L.tam * 1.2 && L.x0 - ant.x0 > L.tam && /[.:;!?]$/.test(ant.texto)) novo = true;
    }
    if (novo) {
      fecha();
      p = { tipo: 'p', y: L.y - L.tam, runs: [], tam: L.tam, centrada, recuo: Math.max(0, L.x0 - esq), x0s: [], x1s: [], linhas: 0 };
    } else {
      const ult = p.runs[p.runs.length - 1];
      if (ult && /[A-Za-zÀ-ÿ]-$/.test(ult.texto) && /^[a-zà-ÿ]/.test(L.texto)) ult.texto = ult.texto.slice(0, -1);
      else p.runs.push(Object.assign({}, runs[0], { texto: ' ' }));
    }
    p.runs.push(...runs);
    p.x0s.push(L.x0); p.x1s.push(L.x1); p.linhas++;
    ant = L;
  }
  fecha();
  for (const b of blocos) {
    if (b.tipo !== 'p') continue;
    b.runs = juntaRuns(b.runs);
    const tamMax = Math.max(...b.runs.map((r) => r.tam));
    const palavras = b.runs.map((r) => r.texto).join('').split(/\s+/).length;
    b.titulo = tamMax >= corpo * 1.45 && palavras <= 25 ? 1 : tamMax >= corpo * 1.18 && palavras <= 30 ? 2 : 0;
    const justif = b.linhas >= 3 && b.x1s.slice(0, -1).every((x) => x > dir - 4) && b.x0s.every((x) => Math.abs(x - b.x0s[0]) < 3 || x < esq + 3);
    const direita = !b.centrada && b.x0s.every((x) => x > esq + largura * 0.4) && b.x1s.every((x) => x > dir - 4);
    b.alin = b.centrada ? 'center' : direita ? 'right' : justif ? 'both' : 'left';
    if (b.alin !== 'left') b.recuo = 0;
    delete b.x0s; delete b.x1s;
  }
  for (const f of figuras || []) blocos.push(f);
  blocos.sort((a, b) => a.y - b.y);
  return { blocos, margemEsq: esq, margemDir: pag.W - dir, corpo };
}

function rPr(r, corpo) {
  let s = '';
  if (r.familia) s += `<w:rFonts w:ascii="${escXml(r.familia)}" w:hAnsi="${escXml(r.familia)}" w:cs="${escXml(r.familia)}"/>`;
  if (r.negrito) s += '<w:b/>';
  if (r.italico) s += '<w:i/>';
  const meio = Math.max(2, Math.round(r.tam * 2));
  if (Math.abs(meio - corpo * 2) >= 1) s += `<w:sz w:val="${meio}"/><w:szCs w:val="${meio}"/>`;
  return s ? `<w:rPr>${s}</w:rPr>` : '';
}
function runXml(r, corpo) {
  return `<w:r>${rPr(r, corpo)}<w:t xml:space="preserve">${escXml(r.texto)}</w:t></w:r>`;
}
const twip = (pt) => Math.max(0, Math.round(pt * 20));
const emu = (pt) => Math.max(1, Math.round(pt * 12700));
function desenhoXml(id, rid, wPt, hPt) {
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${emu(wPt)}" cy="${emu(hPt)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Imagem ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="imagem${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(wPt)}" cy="${emu(hPt)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

/* modelo: { pagina:{W,H}, margens:{t,r,b,l}, corpo, familia, secoes:[{blocos, quebraAntes}], midias:[{bytes, ext}] } */
export async function geraDocx(modelo) {
  const corpo = modelo.corpo || 11;
  const W = modelo.pagina.W, H = modelo.pagina.H;
  const mg = modelo.margens;
  const util = W - mg.l - mg.r;
  const rels = [];
  let idImg = 0;
  let corpoXml = '';
  modelo.secoes.forEach((sec, si) => {
    let primeira = true;
    for (const b of sec.blocos) {
      const quebra = si > 0 && primeira ? '<w:r><w:br w:type="page"/></w:r>' : '';
      primeira = false;
      if (b.tipo === 'p') {
        let ppr = '';
        if (b.titulo) ppr += `<w:pStyle w:val="Heading${b.titulo}"/>`;
        if (b.alin && b.alin !== 'left') ppr += `<w:jc w:val="${b.alin}"/>`;
        if (b.recuo > 6 && b.recuo < util * 0.6) ppr += `<w:ind w:left="${twip(b.recuo)}"/>`;
        corpoXml += `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${quebra}${b.runs.map((r) => runXml(r, b.titulo ? -1 : corpo)).join('')}</w:p>`;
      } else if (b.tipo === 'tabela') {
        const total = b.larguras.reduce((a, c) => a + c, 0);
        const k = total > util ? util / total : 1;
        const cols = b.larguras.map((w) => twip(w * k));
        if (quebra) corpoXml += `<w:p>${quebra}</w:p>`;
        corpoXml += `<w:tbl><w:tblPr><w:tblStyle w:val="TabelaGrade"/><w:tblW w:w="${cols.reduce((a, c) => a + c, 0)}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>${cols.map((c) => `<w:gridCol w:w="${c}"/>`).join('')}</w:tblGrid>`;
        for (const linha of b.grade) {
          corpoXml += '<w:tr>';
          linha.forEach((runs, c) => {
            corpoXml += `<w:tc><w:tcPr><w:tcW w:w="${cols[c]}" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>${(runs || []).map((r) => runXml(r, corpo)).join('')}</w:p></w:tc>`;
          });
          corpoXml += '</w:tr>';
        }
        corpoXml += '</w:tbl><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>';
      } else if (b.tipo === 'img') {
        const m = modelo.midias[b.midia];
        if (!m) continue;
        idImg++;
        const rid = 'rIdImg' + idImg;
        rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/imagem${b.midia + 1}.${m.ext}"/>`);
        let w = b.w, h = b.h;
        const maxW = b.cheia ? W : util, maxH = b.cheia ? H : H - mg.t - mg.b - 4;
        const k = Math.min(1, maxW / w, maxH / h);
        w *= k; h *= k;
        const ppr = b.cheia ? '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>' : (b.alin === 'center' ? '<w:pPr><w:jc w:val="center"/></w:pPr>' : '');
        corpoXml += `<w:p>${ppr}${quebra}${desenhoXml(idImg, rid, w, h)}</w:p>`;
      }
    }
  });
  const orient = W > H ? ' w:orient="landscape"' : '';
  const sect = `<w:sectPr><w:pgSz w:w="${twip(W)}" w:h="${twip(H)}"${orient}/><w:pgMar w:top="${twip(mg.t)}" w:right="${twip(mg.r)}" w:bottom="${twip(mg.b)}" w:left="${twip(mg.l)}" w:header="${twip(Math.min(36, mg.t))}" w:footer="${twip(Math.min(36, mg.b))}" w:gutter="0"/></w:sectPr>`;
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${corpoXml || '<w:p/>'}${sect}</w:body></w:document>`;
  const fam = escXml(modelo.familia || 'Calibri');
  const meio = Math.round(corpo * 2);
  const estilos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${fam}" w:hAnsi="${fam}" w:eastAsia="${fam}" w:cs="${fam}"/><w:sz w:val="${meio}"/><w:szCs w:val="${meio}"/><w:lang w:val="pt-BR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="${modelo.semEspaco ? 0 : 120}" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="${Math.round(meio * 1.6)}"/><w:szCs w:val="${Math.round(meio * 1.6)}"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="${Math.round(meio * 1.3)}"/><w:szCs w:val="${Math.round(meio * 1.3)}"/></w:rPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TabelaGrade"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders></w:tblPr></w:style></w:styles>`;
  const exts = new Set(modelo.midias.map((m) => m.ext));
  const itens = [
    { nome: '[Content_Types].xml', dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${exts.has('png') ? '<Default Extension="png" ContentType="image/png"/>' : ''}${exts.has('jpeg') ? '<Default Extension="jpeg" ContentType="image/jpeg"/>' : ''}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>` },
    { nome: '_rels/.rels', dados: RELS_RAIZ('word/document.xml') },
    { nome: 'docProps/core.xml', dados: CORE() },
    { nome: 'docProps/app.xml', dados: APP('Microsoft Office Word') },
    { nome: 'word/document.xml', dados: doc },
    { nome: 'word/styles.xml', dados: estilos },
    { nome: 'word/_rels/document.xml.rels', dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEstilos" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${rels.join('')}</Relationships>` },
    ...modelo.midias.map((m, i) => ({ nome: `word/media/imagem${i + 1}.${m.ext}`, dados: m.bytes, comprimir: false })),
  ];
  return criaZip(itens, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}

/* ================================================================
   PDF -> PowerPoint (cada página vira um slide com a imagem da página)
   ================================================================ */
export async function geraPptx(slides, larguraPt, alturaPt) {
  /* PowerPoint aceita slides entre 1 e 56 polegadas */
  let cx = Math.round(larguraPt * 12700), cy = Math.round(alturaPt * 12700);
  const k = Math.min(1, 51206400 / Math.max(cx, cy));
  cx = Math.max(914400, Math.round(cx * k)); cy = Math.max(914400, Math.round(cy * k));
  const itens = [];
  const ct = [];
  slides.forEach((s, i) => {
    const n = i + 1;
    /* cada imagem cabe inteira no slide, centralizada */
    const kk = Math.min(cx / s.w, cy / s.h);
    const w = Math.round(s.w * kk), h = Math.round(s.h * kk);
    const x = Math.round((cx - w) / 2), y = Math.round((cy - h) / 2);
    itens.push({ nome: `ppt/slides/slide${n}.xml`, dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="2" name="Página ${n}" descr="Página ${n}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>` });
    itens.push({ nome: `ppt/slides/_rels/slide${n}.xml.rels`, dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/pagina${n}.jpeg"/></Relationships>` });
    itens.push({ nome: `ppt/media/pagina${n}.jpeg`, dados: s.bytes, comprimir: false });
    ct.push(`<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
  });
  const pres = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${10 + i}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${cx}" cy="${cy}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const presRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>${slides.map((_, i) => `<Relationship Id="rId${10 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('')}</Relationships>`;
  const layout = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Em branco"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const base = [
    { nome: '[Content_Types].xml', dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>${ct.join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>` },
    { nome: '_rels/.rels', dados: RELS_RAIZ('ppt/presentation.xml') },
    { nome: 'docProps/core.xml', dados: CORE() },
    { nome: 'docProps/app.xml', dados: APP('Microsoft Office PowerPoint') },
    { nome: 'ppt/presentation.xml', dados: pres },
    { nome: 'ppt/_rels/presentation.xml.rels', dados: presRels },
    { nome: 'ppt/presProps.xml', dados: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>' },
    { nome: 'ppt/viewProps.xml', dados: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>' },
    { nome: 'ppt/tableStyles.xml', dados: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>' },
    { nome: 'ppt/theme/theme1.xml', dados: TEMA },
    { nome: 'ppt/slideMasters/slideMaster1.xml', dados: MESTRE },
    { nome: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${R}theme" Target="../theme/theme1.xml"/></Relationships>` },
    { nome: 'ppt/slideLayouts/slideLayout1.xml', dados: layout },
    { nome: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', dados: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>` },
  ];
  return criaZip(base.concat(itens), 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
}

/* ================================================================
   PDF -> texto
   ================================================================ */
export function paginaParaTexto(pag) {
  const linhas = montaLinhas(pag.itens).concat(linhasGiradas(pag.girados));
  let out = '', ant = null;
  for (const L of linhas) {
    if (ant && L.y - ant.y > 1.7 * Math.max(ant.tam, L.tam) && L.y < 1e8) out += '\n';
    out += L.segs.map((s) => s.texto).join('    ') + '\n';
    ant = L;
  }
  return out;
}

/* ================================================================
   Fluxos completos (usados pela interface e pelos testes)
   ================================================================ */
async function renderiza(page, escala, maxPixels) {
  let vp = page.getViewport({ scale: escala });
  const area = vp.width * vp.height;
  if (area > (maxPixels || 16e6)) vp = page.getViewport({ scale: escala * Math.sqrt((maxPixels || 16e6) / area) });
  const c = document.createElement('canvas');
  c.width = Math.max(1, Math.floor(vp.width)); c.height = Math.max(1, Math.floor(vp.height));
  const ctx = c.getContext('2d', { alpha: false });
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height);
  /* intent 'print': o desenho continua mesmo com a aba em segundo plano */
  await page.render({ canvasContext: ctx, viewport: vp, intent: 'print' }).promise;
  return c;
}
async function jpeg(c, q) {
  const blob = await new Promise((ok) => c.toBlob(ok, 'image/jpeg', q || 0.85));
  return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
}
const limita = (v, a, b) => Math.max(a, Math.min(b, v));

/* op.modo: 'editavel' (texto, tabelas e imagens) ou 'visual' (cada página como imagem) */
export async function pdfParaWord(pdfjs, doc, op, aoProgredir) {
  const n = doc.numPages;
  const midias = [], secoes = [];
  let W = 595.28, H = 841.89, margens = null, corpo = 11, escaneadas = 0;
  const familias = new Map();
  for (let i = 1; i <= n; i++) {
    aoProgredir && aoProgredir(i - 1, n);
    const page = await doc.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    if (i === 1) { W = vp.width; H = vp.height; }
    if (op.modo === 'visual') {
      const c = await renderiza(page, 150 / 72, 12e6);
      const b = await jpeg(c, 0.85); c.width = 0; c.height = 0;
      midias.push({ bytes: b, ext: 'jpeg' });
      secoes.push({ blocos: [{ tipo: 'img', midia: midias.length - 1, w: vp.width, h: vp.height - 1, cheia: true }] });
      page.cleanup();
      continue;
    }
    const pag = await extraiPagina(pdfjs, page, { fontes: true, imagens: true });
    const figuras = [];
    if (pag.escaneada) {
      escaneadas++;
      const c = await renderiza(page, 150 / 72, 12e6);
      const b = await jpeg(c, 0.85); c.width = 0; c.height = 0;
      midias.push({ bytes: b, ext: 'jpeg' });
      figuras.push({ tipo: 'img', midia: midias.length - 1, y: 0, w: vp.width, h: vp.height, alin: 'center' });
    } else {
      for (const im of pag.imagens) {
        const c = await canvasDaImagem(page, im);
        if (!c) continue;
        const r = await canvasParaBytes(c, 2400);
        c.width = 0; c.height = 0;
        if (!r) continue;
        midias.push({ bytes: r.bytes, ext: r.ext });
        const meio = (im.caixa.x0 + im.caixa.x1) / 2;
        figuras.push({ tipo: 'img', midia: midias.length - 1, y: im.caixa.y0, w: im.caixa.x1 - im.caixa.x0, h: im.caixa.y1 - im.caixa.y0, alin: Math.abs(meio - vp.width / 2) < vp.width * 0.08 ? 'center' : 'left' });
      }
    }
    const r = paginaParaBlocos(pag, figuras);
    for (const L of pag.girados) r.blocos.push({ tipo: 'p', y: 1e9, runs: [{ texto: L.s.trim(), negrito: L.negrito, italico: L.italico, tam: L.tam, familia: L.familia }], alin: 'left', recuo: 0, titulo: 0 });
    if (!margens && pag.itens.length) margens = { l: limita(r.margemEsq, 18, 108), r: limita(r.margemDir, 18, 108) };
    if (i <= 5) { corpo = r.corpo || corpo; for (const it of pag.itens) if (it.familia) familias.set(it.familia, (familias.get(it.familia) || 0) + it.s.length); }
    secoes.push({ blocos: r.blocos });
    page.cleanup();
    if (i % 5 === 0) await pausa();
  }
  aoProgredir && aoProgredir(n, n);
  let familia = 'Calibri', mv = 0;
  for (const [k, v] of familias) if (v > mv) { mv = v; familia = k; }
  const visual = op.modo === 'visual';
  const modelo = {
    pagina: { W, H }, corpo, familia, secoes, midias, semEspaco: visual,
    margens: visual ? { t: 0, r: 0, b: 0, l: 0 } : { t: 50, b: 50, l: (margens && margens.l) || 56, r: (margens && margens.r) || 56 },
  };
  const blob = await geraDocx(modelo);
  return { blob, paginas: n, escaneadas };
}

/* op.modo: 'paginas' (uma planilha por página) ou 'unica' (tudo numa planilha) */
export async function pdfParaExcel(pdfjs, doc, op, aoProgredir) {
  const n = doc.numPages;
  const pags = [];
  const textos = [];
  let escaneadas = 0;
  for (let i = 1; i <= n; i++) {
    aoProgredir && aoProgredir(i - 1, n * 1.2);
    const page = await doc.getPage(i);
    const pag = await extraiPagina(pdfjs, page, { fontes: true });
    if (pag.letras < 8) escaneadas++;
    for (const it of pag.itens) if (textos.length < 5000) textos.push(it.s);
    pags.push(pag);
    page.cleanup();
    if (i % 5 === 0) await pausa();
  }
  const dec = separadorDecimal(textos);
  const planilhas = [];
  let celulas = 0;
  if (op.modo === 'unica') {
    const grade = [];
    pags.forEach((pag, i) => { if (i && grade.length) grade.push([]); for (const l of paginaParaPlanilha(pag, dec)) grade.push(l); });
    planilhas.push({ nome: op.nomePlanilha || 'PDF', grade });
  } else pags.forEach((pag, i) => planilhas.push({ nome: (op.nomePagina || 'Página') + ' ' + (i + 1), grade: paginaParaPlanilha(pag, dec) }));
  for (const p of planilhas) for (const l of p.grade) for (const c of l) if (c) celulas++;
  const blob = await geraXlsx(planilhas);
  aoProgredir && aoProgredir(1, 1);
  return { blob, paginas: n, escaneadas, celulas, planilhas: planilhas.length };
}

export async function pdfParaPowerPoint(pdfjs, doc, op, aoProgredir) {
  const n = doc.numPages;
  const slides = [];
  let W = 720, H = 540;
  for (let i = 1; i <= n; i++) {
    aoProgredir && aoProgredir(i - 1, n);
    const page = await doc.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    if (i === 1) { W = vp.width; H = vp.height; }
    const escala = Math.min(op.qualidade === 'alta' ? 220 / 72 : 150 / 72, 2400 / Math.max(vp.width, vp.height) * 1.0);
    const c = await renderiza(page, escala, 12e6);
    const b = await jpeg(c, op.qualidade === 'alta' ? 0.9 : 0.82);
    slides.push({ bytes: b, w: c.width, h: c.height });
    c.width = 0; c.height = 0;
    page.cleanup();
  }
  aoProgredir && aoProgredir(n, n);
  const blob = await geraPptx(slides, W, H);
  return { blob, paginas: n };
}

export async function pdfParaTexto(pdfjs, doc, op, aoProgredir) {
  const n = doc.numPages;
  const partes = [];
  let letras = 0, escaneadas = 0;
  for (let i = 1; i <= n; i++) {
    aoProgredir && aoProgredir(i - 1, n);
    const page = await doc.getPage(i);
    const pag = await extraiPagina(pdfjs, page, {});
    if (pag.letras < 8) escaneadas++;
    letras += pag.letras;
    let t = paginaParaTexto(pag).replace(/\s+$/, '');
    if (op.separar) t = (op.rotuloPagina || '--- Página') + ' ' + i + ' ---\n\n' + t;
    partes.push(t);
    page.cleanup();
    if (i % 10 === 0) await pausa();
  }
  aoProgredir && aoProgredir(n, n);
  const texto = partes.join(op.separar ? '\n\n\n' : '\n\n') + '\n';
  const final = op.quebraWindows ? texto.replace(/\n/g, '\r\n') : texto;
  return { blob: new Blob(['﻿' + final], { type: 'text/plain;charset=utf-8' }), paginas: n, escaneadas, letras, texto };
}
