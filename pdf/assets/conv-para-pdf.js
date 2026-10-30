/* GIF Local, conversões para PDF: leitura de Word (.docx), Excel (.xlsx,
   .csv) e PowerPoint (.pptx) no navegador.
   Cada leitor transforma o arquivo num "modelo" simples (parágrafos,
   tabelas, células, formas e imagens). O desenho do PDF a partir do modelo
   acontece no worker (conv-worker.js), com a biblioteca pdf-lib. */
import { abreZip, leRelacoes, filhos, filho, caminho, descendentes, atr, ErroConversao, imagemParaPdf, pausa } from './conv-base.js?v=3';

/* ---------------- utilidades ---------------- */
const SERIFADAS = /times|cambria|georgia|garamond|book ?antiqua|palatino|constantia|serif|bookman|century|baskerville|didot|caslon|minion|bodoni|rockwell/i;
const MONO = /courier|consolas|lucida console|mono|menlo|monaco|code/i;
export function familiaPdf(nome) {
  if (!nome) return null;
  if (MONO.test(nome)) return 'courier';
  if (SERIFADAS.test(nome) && !/sans/i.test(nome)) return 'times';
  return 'helv';
}
function hexParaRgb(h) {
  if (!h || !/^[0-9a-f]{6}$/i.test(h)) return null;
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}
function rgbParaHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (mx + mn) / 2;
  if (mx !== mn) {
    const d = mx - mn;
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}
function hslParaRgb([h, s, l]) {
  if (s === 0) { const v = Math.round(l * 255); return [v, v, v]; }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const f = (t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  return [Math.round(f(h + 1 / 3) * 255), Math.round(f(h) * 255), Math.round(f(h - 1 / 3) * 255)];
}
function aplicaTint(rgb, tint) {
  if (!tint) return rgb;
  const [h, s, l] = rgbParaHsl(rgb);
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return hslParaRgb([h, s, Math.max(0, Math.min(1, nl))]);
}
const num = (v, pad) => { const n = parseFloat(v); return isFinite(n) ? n : pad; };
function liga(el) { // propriedades liga/desliga do Word: <w:b/>, <w:b w:val="0"/>
  if (!el) return undefined;
  const v = atr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}
async function imagensDoPacote(zip, caminhos, contador) {
  const mapa = {};
  for (const c of caminhos) {
    if (!c || mapa[c] !== undefined) continue;
    const b = await zip.bytes(c);
    const conv = b ? await imagemParaPdf(b) : null;
    if (conv) mapa[c] = conv; else { mapa[c] = null; contador.imagens++; }
  }
  for (const k of Object.keys(mapa)) if (!mapa[k]) delete mapa[k];
  return mapa;
}

/* ================================================================
   WORD (.docx)
   ================================================================ */
function lePPr(el) {
  const p = {};
  if (!el) return p;
  const st = filho(el, 'pStyle'); if (st) p.estilo = atr(st, 'val');
  const jc = filho(el, 'jc');
  if (jc) { const v = atr(jc, 'val'); p.alin = v === 'center' ? 'center' : (v === 'right' || v === 'end') ? 'right' : (v === 'both' || v === 'distribute') ? 'both' : 'left'; }
  const ind = filho(el, 'ind');
  if (ind) {
    const e = atr(ind, 'left') ?? atr(ind, 'start'); if (e != null) p.esq = num(e, 0) / 20;
    const d = atr(ind, 'right') ?? atr(ind, 'end'); if (d != null) p.dir = num(d, 0) / 20;
    const f = atr(ind, 'firstLine'), h = atr(ind, 'hanging');
    if (h != null) p.primeira = -num(h, 0) / 20; else if (f != null) p.primeira = num(f, 0) / 20;
  }
  const sp = filho(el, 'spacing');
  if (sp) {
    if (atr(sp, 'before') != null) p.antes = num(atr(sp, 'before'), 0) / 20;
    if (atr(sp, 'after') != null) p.depois = num(atr(sp, 'after'), 0) / 20;
    if (atr(sp, 'beforeAutospacing') === '1') p.antes = 14;
    if (atr(sp, 'afterAutospacing') === '1') p.depois = 14;
    const ln = atr(sp, 'line');
    if (ln != null) { const r = atr(sp, 'lineRule') || 'auto'; p.linha = r === 'auto' ? { tipo: 'mult', v: num(ln, 240) / 240 } : { tipo: r === 'exact' ? 'exato' : 'min', v: num(ln, 240) / 20 }; }
  }
  const np = filho(el, 'numPr');
  if (np) { const id = filho(np, 'numId'), lv = filho(np, 'ilvl'); if (id) p.numId = atr(id, 'val'); if (lv) p.ilvl = parseInt(atr(lv, 'val'), 10) || 0; }
  if (liga(filho(el, 'pageBreakBefore'))) p.quebraAntes = true;
  const cs = liga(filho(el, 'contextualSpacing')); if (cs !== undefined) p.contextual = cs;
  const shd = filho(el, 'shd'); if (shd) { const c = hexParaRgb(atr(shd, 'fill')); if (c) p.fundo = c; }
  const rp = filho(el, 'rPr'); if (rp) p.rMarca = leRPr(rp);
  return p;
}
function leRPr(el, tema) {
  const r = {};
  if (!el) return r;
  const st = filho(el, 'rStyle'); if (st) r.estilo = atr(st, 'val');
  const b = liga(filho(el, 'b')); if (b !== undefined) r.b = b;
  const i = liga(filho(el, 'i')); if (i !== undefined) r.i = i;
  const u = filho(el, 'u'); if (u) r.u = atr(u, 'val') !== 'none';
  const s = liga(filho(el, 'strike')); if (s !== undefined) r.riscado = s;
  const c = liga(filho(el, 'caps')); if (c !== undefined) r.caps = c;
  const v = liga(filho(el, 'vanish')); if (v !== undefined) r.oculto = v;
  const sz = filho(el, 'sz'); if (sz) r.tam = num(atr(sz, 'val'), 22) / 2;
  const cor = filho(el, 'color'); if (cor) { const h = atr(cor, 'val'); r.cor = h === 'auto' ? null : hexParaRgb(h); }
  const va = filho(el, 'vertAlign'); if (va) r.vert = atr(va, 'val');
  const f = filho(el, 'rFonts');
  if (f) {
    let n = atr(f, 'ascii') || atr(f, 'hAnsi');
    const t = atr(f, 'asciiTheme') || atr(f, 'hAnsiTheme');
    if (!n && t && tema) n = /major/i.test(t) ? tema.maior : tema.menor;
    if (n) r.fonte = n;
  }
  const hl = filho(el, 'highlight');
  if (hl) { const m = { yellow: [255, 255, 0], green: [0, 255, 0], cyan: [0, 255, 255], magenta: [255, 0, 255], blue: [0, 0, 255], red: [255, 0, 0], darkBlue: [0, 0, 139], lightGray: [211, 211, 211], darkGray: [169, 169, 169] }; const x = m[atr(hl, 'val')]; if (x) r.realce = x; }
  return r;
}
async function temaDoc(zip, rels) {
  const t = { maior: 'Calibri Light', menor: 'Calibri' };
  for (const [, r] of rels) if (r.tipo === 'theme') {
    const d = await zip.xml(r.alvo);
    if (!d) break;
    const fs = descendentes(d.documentElement, 'fontScheme')[0];
    const mj = fs && caminho(fs, 'majorFont', 'latin'), mn = fs && caminho(fs, 'minorFont', 'latin');
    if (mj) t.maior = atr(mj, 'typeface') || t.maior;
    if (mn) t.menor = atr(mn, 'typeface') || t.menor;
  }
  return t;
}

class LeitorDocx {
  constructor(zip, tema) { this.zip = zip; this.tema = tema; this.estilos = new Map(); this.padraoP = null; this.padraoT = null; this.defP = {}; this.defR = {}; this.listas = { abs: new Map(), num: new Map() }; this.contadores = new Map(); this.imagens = new Set(); this.aviso = { imagens: 0, objetos: 0 }; this.idsImagem = 0; }

  async carregaEstilos(caminhoEstilos) {
    const d = caminhoEstilos && await this.zip.xml(caminhoEstilos);
    if (!d) return;
    const raiz = d.documentElement;
    const dd = filho(raiz, 'docDefaults');
    if (dd) { this.defR = leRPr(caminho(dd, 'rPrDefault', 'rPr'), this.tema); this.defP = lePPr(caminho(dd, 'pPrDefault', 'pPr')); }
    for (const s of filhos(raiz, 'style')) {
      const id = atr(s, 'styleId'), tipo = atr(s, 'type');
      const base = filho(s, 'basedOn');
      const nome = filho(s, 'name');
      const e = { tipo, base: base ? atr(base, 'val') : null, nome: nome ? atr(nome, 'val') : id, pPr: lePPr(filho(s, 'pPr')), rPr: leRPr(filho(s, 'rPr'), this.tema), tblPr: filho(s, 'tblPr') };
      this.estilos.set(id, e);
      if (atr(s, 'default') === '1' && tipo === 'paragraph') this.padraoP = id;
      if (atr(s, 'default') === '1' && tipo === 'table') this.padraoT = id;
    }
  }
  cadeia(id) {
    const out = [];
    const vistos = new Set();
    while (id && this.estilos.has(id) && !vistos.has(id)) { vistos.add(id); const e = this.estilos.get(id); out.unshift(e); id = e.base; }
    return out;
  }
  async carregaListas(caminhoNum) {
    const d = caminhoNum && await this.zip.xml(caminhoNum);
    if (!d) return;
    for (const a of filhos(d.documentElement, 'abstractNum')) {
      const niveis = [];
      for (const l of filhos(a, 'lvl')) {
        const il = parseInt(atr(l, 'ilvl'), 10) || 0;
        const fmt = filho(l, 'numFmt'), txt = filho(l, 'lvlText'), ini = filho(l, 'start');
        niveis[il] = { fmt: fmt ? atr(fmt, 'val') : 'decimal', texto: txt ? atr(txt, 'val') : '', inicio: ini ? parseInt(atr(ini, 'val'), 10) : 1, pPr: lePPr(filho(l, 'pPr')), rPr: leRPr(filho(l, 'rPr'), this.tema) };
      }
      this.listas.abs.set(atr(a, 'abstractNumId'), niveis);
    }
    for (const n of filhos(d.documentElement, 'num')) {
      const a = filho(n, 'abstractNumId');
      const sobre = {};
      for (const o of filhos(n, 'lvlOverride')) { const so = filho(o, 'startOverride'); if (so) sobre[parseInt(atr(o, 'ilvl'), 10) || 0] = parseInt(atr(so, 'val'), 10); }
      this.listas.num.set(atr(n, 'numId'), { abs: a ? atr(a, 'val') : null, sobre });
    }
  }
  propsParagrafo(pPr) {
    const est = pPr.estilo || this.padraoP;
    const props = Object.assign({}, this.defP);
    const r = Object.assign({}, this.defR);
    for (const e of this.cadeia(est)) { Object.assign(props, e.pPr); Object.assign(r, e.rPr); }
    Object.assign(props, pPr);
    props.estiloId = est || '';
    delete props.estilo;
    props.nomeEstilo = est && this.estilos.has(est) ? this.estilos.get(est).nome : '';
    return { props, rBase: r };
  }
  marcador(props, rBase) {
    if (!props.numId || props.numId === '0') return null;
    const n = this.listas.num.get(props.numId);
    if (!n) return null;
    const niveis = this.listas.abs.get(n.abs) || [];
    const lv = props.ilvl || 0;
    const nv = niveis[lv];
    if (!nv) return null;
    let cont = this.contadores.get(props.numId);
    if (!cont) { cont = []; this.contadores.set(props.numId, cont); }
    for (let k = 0; k <= lv; k++) if (cont[k] == null) cont[k] = ((n.sobre[k] != null ? n.sobre[k] : (niveis[k] && niveis[k].inicio) || 1)) - 1;
    cont[lv]++;
    cont.length = lv + 1;
    const recuo = nv.pPr || {};
    let texto;
    if (nv.fmt === 'bullet') {
      const c = nv.texto || '•';
      texto = /[-§·]/.test(c) || !/^[o\-–•*>]$/.test(c) ? '•' : c;
    } else if (nv.fmt === 'none') texto = '';
    else texto = (nv.texto || '%' + (lv + 1) + '.').replace(/%(\d)/g, (_, k) => { const i = parseInt(k, 10) - 1; const nvk = niveis[i] || nv; return formataContador(cont[i] != null ? cont[i] : 1, nvk.fmt); });
    return { texto, esq: recuo.esq, primeira: recuo.primeira, r: Object.assign({}, rBase, nv.rPr, { fonte: null }) };
  }
  estiloRun(rBase, rPr) {
    const r = Object.assign({}, rBase);
    if (rPr.estilo) for (const e of this.cadeia(rPr.estilo)) Object.assign(r, e.rPr);
    Object.assign(r, rPr);
    return r;
  }

  /* ----- corpo ----- */
  async blocos(el, rels, saida) {
    for (const c of filhos(el)) {
      try {
        const n = c.localName;
        if (n === 'p') await this.paragrafo(c, rels, saida);
        else if (n === 'tbl') saida.push(await this.tabela(c, rels));
        else if (n === 'sdt') await this.blocos(filho(c, 'sdtContent'), rels, saida);
        else if (n === 'customXml' || n === 'ins' || n === 'smartTag') await this.blocos(c, rels, saida);
        else if (n === 'AlternateContent') await this.blocos(filho(c, 'Choice') || filho(c, 'Fallback'), rels, saida);
        else if (n === 'sectPr') saida.push({ tipo: 'secao', sect: c });
      } catch (e) { if (e instanceof ErroConversao) throw e; this.aviso.objetos++; }
    }
  }
  async paragrafo(el, rels, saida) {
    const pPrEl = filho(el, 'pPr');
    const pPr = lePPr(pPrEl);
    const { props, rBase } = this.propsParagrafo(pPr);
    const p = { tipo: 'p', alin: props.alin || 'left', esq: props.esq || 0, dir: props.dir || 0, primeira: props.primeira || 0, antes: props.antes ?? 0, depois: props.depois ?? 0, linha: props.linha || { tipo: 'mult', v: 1 }, runs: [], fundo: props.fundo || null, quebraAntes: !!props.quebraAntes };
    if (/^(heading|título|titulo|title)/i.test(props.nomeEstilo || '')) p.manterProximo = true;
    p.estiloId = props.estiloId; p.contextual = !!props.contextual;
    const m = this.marcador(props, rBase);
    if (m) {
      p.marcador = { texto: m.texto, r: this.runFinal(m.r) };
      if (m.esq != null && pPr.esq == null) p.esq = m.esq;
      if (m.primeira != null && pPr.primeira == null) p.primeira = m.primeira;
    }
    const extras = [];
    this.campo = null;
    await this.conteudo(el, rels, rBase, p.runs, extras);
    p.tamPadrao = this.runFinal(rBase).tam;
    saida.push(p);
    for (const x of extras) saida.push(x);
    const sp = pPrEl && filho(pPrEl, 'sectPr');
    if (sp) saida.push({ tipo: 'secao', sect: sp });
  }
  runFinal(r) {
    return { b: !!r.b, i: !!r.i, u: !!r.u, riscado: !!r.riscado, tam: r.tam || 11, cor: r.cor || null, fonte: familiaPdf(r.fonte) || 'helv', vert: r.vert || null, caps: !!r.caps, realce: r.realce || null };
  }
  async conteudo(el, rels, rBase, runs, extras) {
    for (const c of filhos(el)) {
      const n = c.localName;
      if (n === 'r') await this.run(c, rels, rBase, runs, extras);
      else if (n === 'hyperlink' || n === 'ins' || n === 'smartTag' || n === 'customXml' || n === 'dir' || n === 'bdo') await this.conteudo(c, rels, rBase, runs, extras);
      else if (n === 'sdt') await this.conteudo(filho(c, 'sdtContent'), rels, rBase, runs, extras);
      else if (n === 'fldSimple') {
        const instr = (atr(c, 'instr') || '').trim().split(/\s+/)[0].toUpperCase();
        if (instr === 'PAGE' || instr === 'NUMPAGES') runs.push(Object.assign(this.runFinal(rBase), { campo: instr, t: '0' }));
        else await this.conteudo(c, rels, rBase, runs, extras);
      } else if (n === 'oMath' || n === 'oMathPara') {
        const t = descendentes(c, 't').map((x) => x.textContent).join('');
        if (t) runs.push(Object.assign(this.runFinal(rBase), { t }));
      }
    }
  }
  async run(el, rels, rBase, runs, extras) {
    const r = this.runFinal(this.estiloRun(rBase, leRPr(filho(el, 'rPr'), this.tema)));
    if (r.oculto) return;
    const add = (t) => { if (this.campo && this.campo.fase !== 'resultado') return; if (this.campo && this.campo.suprimir) return; runs.push(Object.assign({}, r, { t: r.caps ? t.toUpperCase() : t })); };
    for (const c of filhos(el)) {
      const n = c.localName;
      if (n === 't') add(c.textContent);
      else if (n === 'tab' || n === 'ptab') { if (!this.campo || this.campo.fase === 'resultado') runs.push(Object.assign({}, r, { t: '', tab: true })); }
      else if (n === 'br' || n === 'cr') {
        const tipo = atr(c, 'type');
        if (tipo === 'page') runs.push({ quebraPagina: true, t: '' });
        else if (tipo !== 'column') runs.push(Object.assign({}, r, { t: '', br: true }));
      } else if (n === 'noBreakHyphen') add('-');
      else if (n === 'sym') add('•');
      else if (n === 'fldChar') {
        const tipo = atr(c, 'fldCharType');
        if (tipo === 'begin') this.campo = { fase: 'instr', instr: '' };
        else if (tipo === 'separate' && this.campo) {
          const k = this.campo.instr.trim().split(/\s+/)[0].toUpperCase();
          this.campo.fase = 'resultado';
          if (k === 'PAGE' || k === 'NUMPAGES') { runs.push(Object.assign({}, r, { campo: k, t: '0' })); this.campo.suprimir = true; }
        } else if (tipo === 'end') this.campo = null;
      } else if (n === 'instrText') { if (this.campo) this.campo.instr += c.textContent; }
      else if (n === 'drawing') await this.desenho(c, rels, runs, extras, r);
      else if (n === 'pict' || n === 'object') await this.vml(c, rels, runs);
      else if (n === 'AlternateContent') {
        const esc = filho(c, 'Choice');
        const alvo = esc && descendentes(esc, 'drawing').length ? esc : (filho(c, 'Fallback') || esc);
        if (alvo) await this.run(alvo, rels, rBase, runs, extras);
      } else if (n === 'footnoteReference' || n === 'endnoteReference') { /* notas não são desenhadas */ }
    }
  }
  async desenho(el, rels, runs, extras, r) {
    const ext = descendentes(el, 'extent')[0];
    const w = ext ? num(atr(ext, 'cx'), 0) / 12700 : 0, h = ext ? num(atr(ext, 'cy'), 0) / 12700 : 0;
    const blip = descendentes(el, 'blip')[0];
    const caixa = descendentes(el, 'txbxContent')[0];
    if (blip) {
      const rel = rels.get(atr(blip, 'embed'));
      if (rel && !rel.externo && w > 0 && h > 0) { this.imagens.add(rel.alvo); runs.push(Object.assign({}, r, { t: '', img: { id: rel.alvo, w, h } })); return; }
    }
    if (caixa) { await this.blocos(caixa, rels, extras); return; }
    if (descendentes(el, 'chart').length || descendentes(el, 'relIds').length) this.aviso.objetos++;
  }
  async vml(el, rels, runs) {
    const im = descendentes(el, 'imagedata')[0];
    if (!im) { if (descendentes(el, 'OLEObject').length) this.aviso.objetos++; return; }
    const rel = rels.get(atr(im, 'id'));
    const forma = im.parentNode;
    const st = (forma && forma.getAttribute && forma.getAttribute('style')) || '';
    const med = (k) => { const m = st.match(new RegExp(k + '\\s*:\\s*([\\d.]+)(pt|in|px|cm|mm)?')); if (!m) return 0; const v = parseFloat(m[1]); return m[2] === 'in' ? v * 72 : m[2] === 'px' ? v * 0.75 : m[2] === 'cm' ? v * 28.35 : m[2] === 'mm' ? v * 2.835 : v; };
    if (rel && !rel.externo) { this.imagens.add(rel.alvo); runs.push({ t: '', img: { id: rel.alvo, w: med('width') || 200, h: med('height') || 150 } }); }
  }
  bordasTabela(tblPr) {
    const lados = (bs) => bs && filhos(bs).some((b) => !['nil', 'none'].includes(atr(b, 'val') || 'none'));
    let res = null;
    const prop = (el) => { const b = el && filho(el, 'tblBorders'); if (b) res = lados(b); };
    const st = filho(tblPr, 'tblStyle');
    for (const e of this.cadeia(st ? atr(st, 'val') : this.padraoT)) prop(e.tblPr);
    prop(tblPr);
    return !!res;
  }
  async tabela(el, rels) {
    const tblPr = filho(el, 'tblPr');
    const grade = filhos(filho(el, 'tblGrid'), 'gridCol').map((g) => num(atr(g, 'w'), 0) / 20);
    let margem = 5.4;
    const st = filho(tblPr, 'tblStyle');
    for (const e of this.cadeia(st ? atr(st, 'val') : this.padraoT)) { const m = e.tblPr && caminho(e.tblPr, 'tblCellMar', 'left'); if (m) margem = num(atr(m, 'w'), 108) / 20; }
    const mm = caminho(tblPr, 'tblCellMar', 'left'); if (mm) margem = num(atr(mm, 'w'), 108) / 20;
    const ind = filho(tblPr, 'tblInd');
    const jc = filho(tblPr, 'jc');
    const t = { tipo: 'tabela', cols: grade, linhas: [], bordas: this.bordasTabela(tblPr), margem, recuo: ind ? num(atr(ind, 'w'), 0) / 20 : 0, alin: jc ? atr(jc, 'val') : 'left' };
    const linhasEl = [];
    const coleta = (pai) => { for (const c of filhos(pai)) { if (c.localName === 'tr') linhasEl.push(c); else if (c.localName === 'sdt') coleta(filho(c, 'sdtContent')); else if (c.localName === 'customXml') coleta(c); } };
    coleta(el);
    for (const tr of linhasEl) {
      const trPr = filho(tr, 'trPr');
      const alt = trPr && filho(trPr, 'trHeight');
      const linha = { alt: alt ? num(atr(alt, 'val'), 0) / 20 : 0, altExata: alt && atr(alt, 'hRule') === 'exact', cabecalho: !!(trPr && liga(filho(trPr, 'tblHeader'))), celulas: [] };
      const gb = trPr && filho(trPr, 'gridBefore');
      if (gb) for (let k = 0; k < (parseInt(atr(gb, 'val'), 10) || 0); k++) linha.celulas.push({ span: 1, vazia: true, blocos: [] });
      const tcs = [];
      const coletaTc = (pai) => { for (const c of filhos(pai)) { if (c.localName === 'tc') tcs.push(c); else if (c.localName === 'sdt') coletaTc(filho(c, 'sdtContent')); else if (c.localName === 'customXml') coletaTc(c); } };
      coletaTc(tr);
      for (const tc of tcs) {
        const tcPr = filho(tc, 'tcPr');
        const gs = tcPr && filho(tcPr, 'gridSpan');
        const vm = tcPr && filho(tcPr, 'vMerge');
        const shd = tcPr && filho(tcPr, 'shd');
        const va = tcPr && filho(tcPr, 'vAlign');
        const cel = { span: gs ? parseInt(atr(gs, 'val'), 10) || 1 : 1, vmerge: vm ? (atr(vm, 'val') === 'restart' ? 'inicio' : 'continua') : null, fundo: shd ? hexParaRgb(atr(shd, 'fill')) : null, valign: va ? atr(va, 'val') : 'top', blocos: [] };
        await this.blocos(tc, rels, cel.blocos);
        cel.blocos = cel.blocos.filter((b) => b.tipo !== 'secao');
        linha.celulas.push(cel);
      }
      t.linhas.push(linha);
    }
    if (!t.cols.length) { const n = Math.max(1, ...t.linhas.map((l) => l.celulas.reduce((a, c) => a + c.span, 0))); t.cols = new Array(n).fill(0); }
    return t;
  }
  leSecao(sp) {
    const s = { w: 595.3, h: 841.9, t: 72, r: 72, b: 72, l: 72, cab: 36, rod: 36, refs: {}, tituloPg: false, tipo: 'nextPage' };
    if (!sp) return s;
    const pg = filho(sp, 'pgSz');
    if (pg) { s.w = num(atr(pg, 'w'), 11906) / 20; s.h = num(atr(pg, 'h'), 16838) / 20; if (atr(pg, 'orient') === 'landscape' && s.w < s.h) [s.w, s.h] = [s.h, s.w]; }
    const mg = filho(sp, 'pgMar');
    if (mg) { s.t = Math.abs(num(atr(mg, 'top'), 1440)) / 20; s.r = num(atr(mg, 'right'), 1440) / 20; s.b = Math.abs(num(atr(mg, 'bottom'), 1440)) / 20; s.l = num(atr(mg, 'left'), 1440) / 20; s.cab = num(atr(mg, 'header'), 720) / 20; s.rod = num(atr(mg, 'footer'), 720) / 20; }
    for (const ref of filhos(sp)) {
      if (ref.localName === 'headerReference' || ref.localName === 'footerReference') s.refs[(ref.localName === 'headerReference' ? 'cab_' : 'rod_') + (atr(ref, 'type') || 'default')] = atr(ref, 'id');
    }
    s.tituloPg = !!liga(filho(sp, 'titlePg'));
    const tp = filho(sp, 'type'); if (tp) s.tipo = atr(tp, 'val');
    return s;
  }
}

export async function leDocx(bytes) {
  const zip = await abreZip(bytes);
  const raiz = await leRelacoes(zip, '');
  let docPath = null;
  for (const [, r] of raiz) if (r.tipo === 'officeDocument') docPath = r.alvo;
  if (!docPath || !zip.tem(docPath)) throw new ErroConversao('NAO_DOCX');
  const rels = await leRelacoes(zip, docPath);
  const tema = await temaDoc(zip, rels);
  const L = new LeitorDocx(zip, tema);
  for (const [, r] of rels) {
    if (r.tipo === 'styles') await L.carregaEstilos(r.alvo);
    if (r.tipo === 'numbering') await L.carregaListas(r.alvo);
  }
  const doc = await zip.xml(docPath);
  if (!doc || doc.documentElement.localName !== 'document') throw new ErroConversao('NAO_DOCX');
  const corpo = descendentes(doc.documentElement, 'body')[0];
  if (!corpo) throw new ErroConversao('NAO_DOCX');
  const blocos = [];
  await L.blocos(corpo, rels, blocos);
  /* separa em seções (cada w:sectPr encerra uma seção) */
  const secoes = [];
  let atual = [];
  for (const b of blocos) {
    if (b.tipo === 'secao') { secoes.push({ blocos: atual, sect: L.leSecao(b.sect) }); atual = []; }
    else atual.push(b);
  }
  if (atual.length || !secoes.length) secoes.push({ blocos: atual, sect: L.leSecao(filho(corpo, 'sectPr')) });
  /* cabeçalhos e rodapés */
  const partes = {};
  for (const s of secoes) {
    for (const [chave, id] of Object.entries(s.sect.refs)) {
      const r = rels.get(id);
      if (!r) continue;
      if (!partes[r.alvo]) {
        const d = await zip.xml(r.alvo);
        const relsParte = await leRelacoes(zip, r.alvo);
        const bl = [];
        if (d) { L.contadores = new Map(); await L.blocos(d.documentElement, relsParte, bl); }
        partes[r.alvo] = bl.filter((b) => b.tipo !== 'secao');
      }
      s[chave] = partes[r.alvo];
    }
    delete s.sect.refs;
  }
  await pausa();
  const imagens = await imagensDoPacote(zip, [...L.imagens], L.aviso);
  return { tipo: 'doc', secoes, imagens, avisos: L.aviso };
}

function formataContador(n, fmt) {
  switch (fmt) {
    case 'lowerLetter': return letras(n).toLowerCase();
    case 'upperLetter': return letras(n);
    case 'lowerRoman': return romano(n).toLowerCase();
    case 'upperRoman': return romano(n);
    case 'decimalZero': return n < 10 ? '0' + n : String(n);
    case 'ordinal': return n + 'º';
    default: return String(n);
  }
}
function letras(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s || 'A'; }
function romano(n) {
  const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = ''; for (const [v, l] of t) while (n >= v) { s += l; n -= v; } return s || 'I';
}

/* ================================================================
   EXCEL (.xlsx e .csv)
   ================================================================ */
const PALETA = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
function corExcel(el, temaCores) {
  if (!el) return null;
  let rgb = null;
  const r = el.getAttribute('rgb');
  if (r && r.length >= 6) rgb = hexParaRgb(r.slice(-6));
  else if (el.getAttribute('theme') != null) { const t = temaCores[parseInt(el.getAttribute('theme'), 10)]; if (t) rgb = t.slice(); }
  else if (el.getAttribute('indexed') != null) { const i = parseInt(el.getAttribute('indexed'), 10); if (i < PALETA.length) rgb = hexParaRgb(PALETA[i]); }
  if (rgb && el.getAttribute('tint')) rgb = aplicaTint(rgb, parseFloat(el.getAttribute('tint')));
  return rgb;
}
const FORMATOS_INTERNOS = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '"R$" #,##0;-"R$" #,##0', 6: '"R$" #,##0;[Red]-"R$" #,##0', 7: '"R$" #,##0.00;-"R$" #,##0.00', 8: '"R$" #,##0.00;[Red]-"R$" #,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??', 14: 'DATA_CURTA', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'DATA_CURTA h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@' };
for (let i = 27; i <= 36; i++) FORMATOS_INTERNOS[i] = 'DATA_CURTA';
for (let i = 50; i <= 58; i++) FORMATOS_INTERNOS[i] = 'DATA_CURTA';

const LOCAIS = {
  pt: { dec: ',', mil: '.', curta: 'dd/mm/yyyy', meses: ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'], dias: ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'], v: 'VERDADEIRO', f: 'FALSO' },
  en: { dec: '.', mil: ',', curta: 'm/d/yyyy', meses: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], dias: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], v: 'TRUE', f: 'FALSE' },
  es: { dec: ',', mil: '.', curta: 'dd/mm/yyyy', meses: ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'], dias: ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'], v: 'VERDADERO', f: 'FALSO' },
  fr: { dec: ',', mil: ' ', curta: 'dd/mm/yyyy', meses: ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'], dias: ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'], v: 'VRAI', f: 'FAUX' },
  de: { dec: ',', mil: '.', curta: 'dd.mm.yyyy', meses: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'], dias: ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'], v: 'WAHR', f: 'FALSCH' },
  it: { dec: ',', mil: '.', curta: 'dd/mm/yyyy', meses: ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno', 'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'], dias: ['domenica', 'lunedì', 'martedì', 'mercoledì', 'giovedì', 'venerdì', 'sabato'], v: 'VERO', f: 'FALSO' },
};
function divideSecoes(cod) {
  const out = []; let atual = '', aspas = false;
  for (let i = 0; i < cod.length; i++) {
    const c = cod[i];
    if (c === '"') aspas = !aspas;
    if (c === '\\' && i + 1 < cod.length) { atual += c + cod[++i]; continue; }
    if (c === ';' && !aspas) { out.push(atual); atual = ''; continue; }
    atual += c;
  }
  out.push(atual);
  return out;
}
function ehFormatoData(cod) {
  const s = cod.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  return /[dyhs]/i.test(s) || /m/i.test(s) && !/[0#?]/.test(s);
}
function formataGeral(v, loc) {
  if (!isFinite(v)) return String(v);
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  let s;
  const a = Math.abs(v);
  if (a !== 0 && (a >= 1e11 || a < 1e-9)) s = v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E');
  else { s = String(parseFloat(v.toPrecision(11))); if (/e/.test(s)) s = v.toFixed(10).replace(/0+$/, '').replace(/\.$/, ''); }
  return s.replace('.', loc.dec);
}
function formataData(v, cod, loc) {
  const ms = Math.round((v - 25569) * 86400000);
  const d = new Date(ms);
  if (isNaN(d.getTime())) return String(v);
  const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), h = d.getUTCHours(), mi = d.getUTCMinutes(), se = d.getUTCSeconds(), dia = d.getUTCDay();
  const ampm = /AM\/PM|A\/P/i.test(cod);
  const toks = cod.match(/"[^"]*"|\\.|\[h+\]|\[m+\]|\[s+\]|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P|\.0+|./gi) || [];
  let out = '';
  const pad = (n) => (n < 10 ? '0' : '') + n;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i], tl = t.toLowerCase();
    const anterior = toks.slice(0, i).reverse().find((x) => /^[hms]/i.test(x) || /^\[h/i.test(x)) || '';
    const proximo = toks.slice(i + 1).find((x) => /^[hmsdy]/i.test(x)) || '';
    if (t[0] === '"') out += t.slice(1, -1);
    else if (t[0] === '\\') out += t[1];
    else if (/^\[h+\]$/i.test(t)) out += Math.floor(v * 24);
    else if (/^\[m+\]$/i.test(t)) out += Math.floor(v * 1440);
    else if (/^\[s+\]$/i.test(t)) out += Math.floor(v * 86400);
    else if (tl === 'yyyy') out += Y;
    else if (tl === 'yy') out += pad(Y % 100);
    else if (tl === 'mmmmm') out += loc.meses[M][0];
    else if (tl === 'mmmm') out += loc.meses[M];
    else if (tl === 'mmm') out += loc.meses[M].slice(0, 3);
    else if (tl === 'mm' || tl === 'm') {
      const minuto = /^h|^\[h/i.test(anterior) || /^s/i.test(proximo);
      out += minuto ? (tl === 'mm' ? pad(mi) : mi) : (tl === 'mm' ? pad(M + 1) : M + 1);
    } else if (tl === 'dddd') out += loc.dias[dia];
    else if (tl === 'ddd') out += loc.dias[dia].slice(0, 3);
    else if (tl === 'dd') out += pad(D);
    else if (tl === 'd') out += D;
    else if (tl === 'hh' || tl === 'h') { const hh = ampm ? (h % 12 || 12) : h; out += tl === 'hh' ? pad(hh) : hh; }
    else if (tl === 'ss') out += pad(se);
    else if (tl === 's') out += se;
    else if (tl === 'am/pm') out += h < 12 ? 'AM' : 'PM';
    else if (tl === 'a/p') out += h < 12 ? 'A' : 'P';
    else if (/^\.0+$/.test(t)) out += loc.dec + String(Math.round((v * 86400 % 1) * Math.pow(10, t.length - 1))).padStart(t.length - 1, '0');
    else if (t === '_' ) { i++; out += ' '; }
    else if (t === '*') i++;
    else out += t;
  }
  return out;
}
function formataNumeroSecao(v, cod, loc) {
  let s = cod;
  let prefixo = '', sufixo = '';
  // literais e moeda [$R$-416]
  s = s.replace(/\[\$([^\]-]*)(-[^\]]*)?\]/g, (_, m) => (m ? '"' + m + '"' : ''));
  s = s.replace(/\[[^\]]*\]/g, '');
  const pct = /%/.test(s.replace(/"[^"]*"/g, ''));
  if (pct) v *= 100;
  const partes = s.match(/"[^"]*"|\\.|_.|\*.|[0#?.,]+(?:[Ee][+\-][0#]+)?|\/|@|./g) || [];
  /* localiza o "núcleo" numérico */
  let nucleo = '', ini = -1, fim = -1;
  for (let i = 0; i < partes.length; i++) if (/^[0#?.,]+$/.test(partes[i]) || /^[0#?.,]+[Ee][+\-][0#]+$/.test(partes[i])) { if (ini < 0) ini = i; fim = i; }
  if (ini >= 0) for (let i = ini; i <= fim; i++) nucleo += /^[0#?.,Ee+\-]+$/.test(partes[i]) ? partes[i] : '';
  const literal = (p) => p[0] === '"' ? p.slice(1, -1) : p[0] === '\\' ? p[1] : p[0] === '_' ? ' ' : p[0] === '*' ? '' : p === '%' ? '%' : p;
  for (let i = 0; i < partes.length; i++) {
    if (ini >= 0 && i >= ini && i <= fim) continue;
    if (ini < 0 || i < ini) prefixo += literal(partes[i]); else sufixo += literal(partes[i]);
  }
  if (!nucleo) return prefixo + sufixo;
  const exp = nucleo.match(/[Ee][+\-]([0#]+)/);
  if (exp) {
    const casas = ((nucleo.split(/[Ee]/)[0].split('.')[1]) || '').replace(/[^0#]/g, '').length;
    let [m, e] = Math.abs(v).toExponential(casas).split('e');
    const ee = parseInt(e, 10);
    return prefixo + (v < 0 ? '-' : '') + m.replace('.', loc.dec) + 'E' + (ee < 0 ? '-' : '+') + String(Math.abs(ee)).padStart(exp[1].length, '0') + sufixo;
  }
  const [ip, dp = ''] = nucleo.split('.');
  const milhar = /[0#?],[0#?]/.test(ip);
  let escala = 0;
  const ipLimpo = ip.replace(/,+$/, (m) => { escala = m.length; return ''; });
  v = v / Math.pow(1000, escala);
  const casasMin = (dp.match(/0/g) || []).length, casasMax = (dp.match(/[0#?]/g) || []).length;
  let txt = Math.abs(v).toFixed(casasMax);
  let [ii, dd = ''] = txt.split('.');
  while (dd.length > casasMin && dd.endsWith('0')) dd = dd.slice(0, -1);
  const minInt = (ipLimpo.match(/0/g) || []).length;
  if (ii === '0' && minInt === 0) ii = '';
  ii = ii.padStart(minInt, '0');
  if (milhar) ii = ii.replace(/\B(?=(\d{3})+(?!\d))/g, loc.mil);
  return prefixo + ii + (dd ? loc.dec + dd : (casasMax && casasMin ? loc.dec : '')) + sufixo;
}
export function formataValor(v, cod, idioma) {
  const loc = LOCAIS[idioma] || LOCAIS.pt;
  if (cod == null || /^general$/i.test(cod) || cod === '') return formataGeral(v, loc);
  cod = cod.replace(/DATA_CURTA/g, loc.curta);
  const secoes = divideSecoes(cod);
  let sec = secoes[0], val = v;
  if (secoes.length > 1) {
    if (v < 0 && secoes[1] !== undefined) { sec = secoes[1]; val = -v; }
    else if (v === 0 && secoes[2] !== undefined) sec = secoes[2];
  }
  const cond = sec.match(/\[(<|>|=|<=|>=|<>)(-?[\d.]+)\]/);
  if (cond) { /* formatos condicionais raros: usa a primeira seção */ }
  if (/^\s*@\s*$/.test(sec)) return String(v);
  if (/^general$/i.test(sec.replace(/\[[^\]]*\]/g, '').trim())) return formataGeral(val, loc);
  if (ehFormatoData(sec)) return formataData(v, sec.replace(/\[(?!h\]|hh\]|m\]|mm\]|s\]|ss\])[^\]]*\]/gi, ''), loc);
  if (/\?\/\?/.test(sec)) return formataGeral(val, loc);
  const res = formataNumeroSecao(val, sec, loc);
  return (v < 0 && secoes.length === 1 ? '-' : '') + res;
}

export async function leXlsx(bytes, idioma) {
  const zip = await abreZip(bytes);
  const raiz = await leRelacoes(zip, '');
  let wb = null;
  for (const [, r] of raiz) if (r.tipo === 'officeDocument') wb = r.alvo;
  if (!wb || !zip.tem(wb)) throw new ErroConversao('NAO_XLSX');
  const rels = await leRelacoes(zip, wb);
  const loc = LOCAIS[idioma] || LOCAIS.pt;
  /* tema (cores) */
  let temaCores = [];
  for (const [, r] of rels) if (r.tipo === 'theme') {
    const d = await zip.xml(r.alvo);
    const cs = d && descendentes(d.documentElement, 'clrScheme')[0];
    if (cs) {
      const pega = (n) => { const e = filho(cs, n); const c = e && filho(e, 'srgbClr'); const s = e && filho(e, 'sysClr'); return hexParaRgb(c ? atr(c, 'val') : s ? atr(s, 'lastClr') : null) || [0, 0, 0]; };
      temaCores = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(pega);
    }
  }
  /* textos compartilhados */
  const textos = [];
  for (const [, r] of rels) if (r.tipo === 'sharedStrings') {
    const d = await zip.xml(r.alvo);
    if (d) for (const si of filhos(d.documentElement, 'si')) textos.push(descendentes(si, 't').filter((t) => !t.parentNode || t.parentNode.localName !== 'rPh').map((t) => t.textContent).join(''));
  }
  /* estilos */
  const xfs = [], fontes = [], fundos = [], bordas = [], fmts = {};
  for (const [, r] of rels) if (r.tipo === 'styles') {
    const d = await zip.xml(r.alvo);
    if (!d) continue;
    const S = d.documentElement;
    for (const f of filhos(filho(S, 'numFmts'), 'numFmt')) fmts[atr(f, 'numFmtId')] = atr(f, 'formatCode');
    for (const f of filhos(filho(S, 'fonts'), 'font')) {
      const sz = filho(f, 'sz'), nm = filho(f, 'name');
      fontes.push({ b: liga(filho(f, 'b')) || false, i: liga(filho(f, 'i')) || false, u: !!filho(f, 'u') && atr(filho(f, 'u'), 'val') !== 'none', tam: sz ? num(atr(sz, 'val'), 11) : 11, cor: corExcel(filho(f, 'color'), temaCores), fonte: familiaPdf(nm ? atr(nm, 'val') : 'Calibri') });
    }
    for (const f of filhos(filho(S, 'fills'), 'fill')) {
      const pf = filho(f, 'patternFill');
      const gf = filho(f, 'gradientFill');
      if (pf && atr(pf, 'patternType') && atr(pf, 'patternType') !== 'none') fundos.push(corExcel(filho(pf, 'fgColor'), temaCores) || [255, 255, 255]);
      else if (gf) { const st = descendentes(gf, 'color')[0]; fundos.push(corExcel(st, temaCores)); }
      else fundos.push(null);
    }
    for (const b of filhos(filho(S, 'borders'), 'border')) {
      const lado = (n) => { const e = filho(b, n); const st = e && atr(e, 'style'); return st && st !== 'none' ? { larg: /thick/.test(st) ? 1.5 : /medium/.test(st) ? 1 : 0.5, cor: corExcel(filho(e, 'color'), temaCores) || [0, 0, 0] } : null; };
      bordas.push({ e: lado('left'), d: lado('right'), t: lado('top'), b: lado('bottom') });
    }
    for (const x of filhos(filho(S, 'cellXfs'), 'xf')) {
      const al = filho(x, 'alignment');
      xfs.push({ fmt: atr(x, 'numFmtId') || '0', fonte: parseInt(atr(x, 'fontId') || '0', 10), fundo: parseInt(atr(x, 'fillId') || '0', 10), borda: parseInt(atr(x, 'borderId') || '0', 10), h: al ? atr(al, 'horizontal') : null, v: al ? atr(al, 'vertical') : null, quebra: al ? atr(al, 'wrapText') === '1' || atr(al, 'wrapText') === 'true' : false, recuo: al ? num(atr(al, 'indent'), 0) : 0 });
    }
  }
  const wbDoc = await zip.xml(wb);
  if (!wbDoc || wbDoc.documentElement.localName !== 'workbook') throw new ErroConversao('NAO_XLSX');
  const d1904 = !!descendentes(wbDoc.documentElement, 'workbookPr').find((w) => atr(w, 'date1904') === '1' || atr(w, 'date1904') === 'true');
  const folhas = [];
  let celulasTotal = 0;
  const fonte0 = fontes[0] || { tam: 11, fonte: 'helv', b: false, i: false, cor: null };
  for (const s of descendentes(wbDoc.documentElement, 'sheet')) {
    const nome = atr(s, 'name');
    const oculta = atr(s, 'state') === 'hidden' || atr(s, 'state') === 'veryHidden';
    const rel = rels.get(atr(s, 'id'));
    if (!rel) continue;
    const d = await zip.xml(rel.alvo);
    if (!d || d.documentElement.localName !== 'worksheet') { folhas.push({ nome, oculta, vazia: true, naoSuportada: true }); continue; }
    const W = d.documentElement;
    const fp = filho(W, 'sheetFormatPr');
    const larguraPadrao = fp && atr(fp, 'defaultColWidth') ? num(atr(fp, 'defaultColWidth'), 8.43) : fp && atr(fp, 'baseColWidth') ? num(atr(fp, 'baseColWidth'), 8) + 0.71 : 8.43;
    const alturaPadrao = fp && atr(fp, 'defaultRowHeight') ? num(atr(fp, 'defaultRowHeight'), 15) : 15;
    const colunas = [];
    for (const c of filhos(filho(W, 'cols'), 'col')) {
      const a = parseInt(atr(c, 'min'), 10), b = Math.min(parseInt(atr(c, 'max'), 10), 16384);
      for (let k = a; k <= Math.min(b, a + 1000); k++) colunas[k - 1] = { w: atr(c, 'width') != null ? num(atr(c, 'width'), larguraPadrao) : larguraPadrao, oculta: atr(c, 'hidden') === '1' || atr(c, 'hidden') === 'true' };
    }
    const linhas = [];
    let maxC = -1;
    for (const row of filhos(filho(W, 'sheetData'), 'row')) {
      const ri = parseInt(atr(row, 'r'), 10) - 1 || linhas.length;
      const lin = { h: atr(row, 'ht') != null ? num(atr(row, 'ht'), alturaPadrao) : null, oculta: atr(row, 'hidden') === '1' || atr(row, 'hidden') === 'true', celulas: [] };
      let ci = -1;
      for (const c of filhos(row, 'c')) {
        const ref = atr(c, 'r');
        if (ref) { const m = ref.match(/^([A-Z]+)/); let n = 0; for (const ch of m[1]) n = n * 26 + ch.charCodeAt(0) - 64; ci = n - 1; } else ci++;
        if (++celulasTotal > 400000) throw new ErroConversao('PLANILHA_GRANDE');
        const st = xfs[parseInt(atr(c, 's') || '0', 10)] || xfs[0] || { fmt: '0', fonte: 0, fundo: 0, borda: 0 };
        const t = atr(c, 't');
        const vEl = filho(c, 'v');
        let texto = '', numero = false;
        if (t === 's') texto = textos[parseInt(vEl ? vEl.textContent : '-1', 10)] ?? '';
        else if (t === 'inlineStr') texto = descendentes(filho(c, 'is'), 't').map((x) => x.textContent).join('');
        else if (t === 'str') texto = vEl ? vEl.textContent : '';
        else if (t === 'b') texto = vEl && vEl.textContent === '1' ? loc.v : loc.f;
        else if (t === 'e') texto = vEl ? vEl.textContent : '#N/A';
        else if (vEl && vEl.textContent !== '') {
          let v = parseFloat(vEl.textContent);
          if (t === 'd') { const ms = Date.parse(vEl.textContent); v = isNaN(ms) ? NaN : ms / 86400000 + 25569; }
          if (isFinite(v)) {
            const cod = fmts[st.fmt] ?? FORMATOS_INTERNOS[st.fmt] ?? 'General';
            if (d1904 && ehFormatoData(cod)) v += 1462;
            try { texto = formataValor(v, cod, idioma); } catch (e) { texto = String(v); }
            numero = true;
          }
        }
        const f = fontes[st.fonte] || fonte0;
        const cel = { c: ci, texto, num: numero, alin: st.h || null, valin: st.v || 'bottom', quebra: st.quebra, recuo: st.recuo, b: f.b, i: f.i, u: f.u, tam: f.tam, cor: f.cor, fonte: f.fonte, fundo: fundos[st.fundo] || null, borda: bordas[st.borda] || null };
        if (!texto && !cel.fundo && !(cel.borda && (cel.borda.e || cel.borda.d || cel.borda.t || cel.borda.b))) continue;
        lin.celulas.push(cel);
        if (texto || cel.fundo) maxC = Math.max(maxC, ci);
        else if (cel.borda) maxC = Math.max(maxC, ci);
      }
      linhas[ri] = lin;
    }
    const mescla = filhos(filho(W, 'mergeCells'), 'mergeCell').map((m) => {
      const [a, b] = (atr(m, 'ref') || '').split(':');
      const p = (x) => { const mm = x && x.match(/^([A-Z]+)(\d+)$/); if (!mm) return null; let n = 0; for (const ch of mm[1]) n = n * 26 + ch.charCodeAt(0) - 64; return [parseInt(mm[2], 10) - 1, n - 1]; };
      const A = p(a), B = p(b || a);
      return A && B ? { r0: A[0], c0: A[1], r1: B[0], c1: B[1] } : null;
    }).filter(Boolean);
    let maxR = linhas.length - 1;
    while (maxR >= 0 && (!linhas[maxR] || !linhas[maxR].celulas.some((c) => c.texto || c.fundo))) maxR--;
    const ps = filho(W, 'pageSetup');
    const po = filho(W, 'printOptions');
    folhas.push({
      nome, oculta, colunas: Array.from({ length: maxC + 1 }, (_, k) => colunas[k] || { w: larguraPadrao, oculta: false }), alturaPadrao,
      linhas: linhas.slice(0, maxR + 1), mescla, maxC, maxR,
      orientacao: ps ? atr(ps, 'orientation') : null, grade: po ? atr(po, 'gridLines') === '1' : null,
    });
    await pausa();
  }
  return { tipo: 'planilhas', folhas, fonteBase: fonte0 };
}

/* CSV: separador detectado (; , tab), texto UTF-8 ou Windows-1252 */
export function leCsv(bytes) {
  let texto;
  try { texto = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
  catch (e) { texto = new TextDecoder('windows-1252').decode(bytes); }
  texto = texto.replace(/^﻿/, '');
  const amostra = texto.slice(0, 5000).split(/\r?\n/).slice(0, 10);
  const conta = (sep) => amostra.reduce((a, l) => a + (l.split(sep).length - 1), 0);
  const sep = [';', ',', '\t', '|'].sort((a, b) => conta(b) - conta(a))[0];
  const linhas = [];
  let campo = '', linha = [], aspas = false;
  for (let i = 0; i < texto.length; i++) {
    const c = texto[i];
    if (aspas) {
      if (c === '"') { if (texto[i + 1] === '"') { campo += '"'; i++; } else aspas = false; }
      else campo += c;
    } else if (c === '"' && campo === '') aspas = true;
    else if (c === sep) { linha.push(campo); campo = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && texto[i + 1] === '\n') i++; linha.push(campo); linhas.push(linha); linha = []; campo = ''; if (linhas.length > 200000) throw new ErroConversao('PLANILHA_GRANDE'); }
    else campo += c;
  }
  if (campo || linha.length) { linha.push(campo); linhas.push(linha); }
  while (linhas.length && linhas[linhas.length - 1].every((x) => !x)) linhas.pop();
  const maxC = Math.max(0, ...linhas.map((l) => l.length)) - 1;
  const larg = [];
  for (const l of linhas) l.forEach((x, k) => { larg[k] = Math.max(larg[k] || 6, Math.min(50, x.length + 1)); });
  const numero = /^-?[\d.,]+%?$/;
  return {
    tipo: 'planilhas', fonteBase: { tam: 10, fonte: 'helv' },
    folhas: [{ nome: 'CSV', colunas: Array.from({ length: maxC + 1 }, (_, k) => ({ w: larg[k] || 8.43, oculta: false })), alturaPadrao: 15, mescla: [], maxC, maxR: linhas.length - 1,
      linhas: linhas.map((l, r) => ({ h: null, celulas: l.map((x, c) => ({ c, texto: x, num: numero.test(x.trim()), b: r === 0, tam: 10, fonte: 'helv', valin: 'bottom' })).filter((c) => c.texto) })) }],
  };
}

/* ================================================================
   POWERPOINT (.pptx)
   ================================================================ */
const EMU = 12700;
function corDrawing(el, ctx) {
  /* el: elemento que contém a cor (solidFill, fontRef...). Devolve { rgb, alfa } */
  if (!el) return null;
  const c = filhos(el).find((x) => /Clr$/.test(x.localName));
  if (!c) return null;
  let rgb = null;
  const n = c.localName, v = atr(c, 'val');
  if (n === 'srgbClr') rgb = hexParaRgb(v);
  else if (n === 'sysClr') rgb = hexParaRgb(atr(c, 'lastClr')) || (v === 'window' ? [255, 255, 255] : [0, 0, 0]);
  else if (n === 'schemeClr') {
    let k = v;
    if (k === 'phClr' && ctx.phClr) rgb = ctx.phClr.slice();
    else { k = ctx.mapa[k] || k; rgb = ctx.tema[k] ? ctx.tema[k].slice() : null; }
  } else if (n === 'prstClr') rgb = { black: [0, 0, 0], white: [255, 255, 255], red: [255, 0, 0], blue: [0, 0, 255], green: [0, 128, 0], yellow: [255, 255, 0], gray: [128, 128, 128] }[v] || [0, 0, 0];
  else if (n === 'scrgbClr') rgb = [num(atr(c, 'r'), 0) / 1000 * 2.55, num(atr(c, 'g'), 0) / 1000 * 2.55, num(atr(c, 'b'), 0) / 1000 * 2.55].map(Math.round);
  else if (n === 'hslClr') rgb = hslParaRgb([num(atr(c, 'hue'), 0) / 21600000, num(atr(c, 'sat'), 0) / 100000, num(atr(c, 'lum'), 0) / 100000]);
  if (!rgb) return null;
  let alfa = 1;
  for (const m of filhos(c)) {
    const val = num(atr(m, 'val'), 0) / 100000;
    if (m.localName === 'alpha') alfa = val;
    else if (m.localName === 'lumMod' || m.localName === 'lumOff') {
      const hsl = rgbParaHsl(rgb);
      hsl[2] = m.localName === 'lumMod' ? hsl[2] * val : hsl[2] + val;
      hsl[2] = Math.max(0, Math.min(1, hsl[2]));
      rgb = hslParaRgb(hsl);
    } else if (m.localName === 'tint') rgb = rgb.map((x) => Math.round(255 - (255 - x) * val));
    else if (m.localName === 'shade') rgb = rgb.map((x) => Math.round(x * val));
  }
  return { rgb, alfa };
}
function preenchimento(spPr, ctx, estilo) {
  if (spPr) {
    if (filho(spPr, 'noFill')) return null;
    const sf = filho(spPr, 'solidFill'); if (sf) return corDrawing(sf, ctx);
    const gf = filho(spPr, 'gradFill');
    if (gf) { const gs = descendentes(gf, 'gs'); if (gs.length) { const a = corDrawing(gs[0], ctx), b = corDrawing(gs[gs.length - 1], ctx); if (a && b) return { rgb: a.rgb.map((x, i) => Math.round((x + b.rgb[i]) / 2)), alfa: a.alfa }; return a || b; } }
    if (filho(spPr, 'blipFill') || filho(spPr, 'pattFill') || filho(spPr, 'grpFill')) return null;
  }
  if (estilo) { const fr = filho(estilo, 'fillRef'); if (fr && atr(fr, 'idx') !== '0') return corDrawing(fr, ctx); }
  return undefined;
}
function contorno(spPr, ctx, estilo) {
  const ln = spPr && filho(spPr, 'ln');
  if (ln && filho(ln, 'noFill')) return null;
  let cor = ln && filho(ln, 'solidFill') ? corDrawing(filho(ln, 'solidFill'), ctx) : null;
  let larg = ln && atr(ln, 'w') != null ? num(atr(ln, 'w'), 12700) / EMU : null;
  if (!cor && estilo) { const lr = filho(estilo, 'lnRef'); if (lr && atr(lr, 'idx') !== '0') { cor = corDrawing(lr, ctx); if (larg == null) larg = 0.75 * parseInt(atr(lr, 'idx'), 10); } }
  if (!cor) return ln && larg != null ? undefined : (ln ? undefined : undefined);
  return { cor, larg: larg == null ? 0.75 : larg };
}
function geometria(el) {
  const x = el && filho(el, 'xfrm');
  if (!x) return null;
  const o = filho(x, 'off'), e = filho(x, 'ext');
  if (!o || !e) return null;
  return { x: num(atr(o, 'x'), 0) / EMU, y: num(atr(o, 'y'), 0) / EMU, w: num(atr(e, 'cx'), 0) / EMU, h: num(atr(e, 'cy'), 0) / EMU, rot: num(atr(x, 'rot'), 0) / 60000, fh: atr(x, 'flipH') === '1', fv: atr(x, 'flipV') === '1', chOff: filho(x, 'chOff'), chExt: filho(x, 'chExt') };
}
/* níveis de estilo de texto (lvl1pPr ... lvl9pPr) */
function leNiveis(lst) {
  const out = [];
  if (!lst) return out;
  for (let i = 1; i <= 9; i++) { const l = filho(lst, 'lvl' + i + 'pPr'); if (l) out[i - 1] = l; }
  return out;
}
function propsNivel(pPr, ctx) {
  const p = {};
  if (!pPr) return p;
  const a = atr(pPr, 'algn'); if (a) p.alin = a === 'ctr' ? 'center' : a === 'r' ? 'right' : (a === 'just' || a === 'dist') ? 'both' : 'left';
  if (atr(pPr, 'marL') != null) p.esq = num(atr(pPr, 'marL'), 0) / EMU;
  if (atr(pPr, 'indent') != null) p.primeira = num(atr(pPr, 'indent'), 0) / EMU;
  const esp = (n) => { const e = filho(pPr, n); if (!e) return undefined; const pt = filho(e, 'spcPts'), pc = filho(e, 'spcPct'); return pt ? { pts: num(atr(pt, 'val'), 0) / 100 } : pc ? { pct: num(atr(pc, 'val'), 0) / 100000 } : undefined; };
  const sb = esp('spcBef'); if (sb) p.antes = sb;
  const sa = esp('spcAft'); if (sa) p.depois = sa;
  const ls = esp('lnSpc'); if (ls) p.linha = ls;
  if (filho(pPr, 'buNone')) p.bullet = null;
  const bc = filho(pPr, 'buChar'); if (bc) p.bullet = { tipo: 'char', c: atr(bc, 'char') || '•' };
  const ba = filho(pPr, 'buAutoNum'); if (ba) p.bullet = { tipo: 'num', esquema: atr(ba, 'type') || 'arabicPeriod', inicio: num(atr(ba, 'startAt'), 1) };
  const bcl = filho(pPr, 'buClr'); if (bcl) p.buCor = corDrawing(bcl, ctx);
  const d = filho(pPr, 'defRPr'); if (d) p.r = propsRun(d, ctx);
  return p;
}
function propsRun(rPr, ctx) {
  const r = {};
  if (!rPr) return r;
  if (atr(rPr, 'sz') != null) r.tam = num(atr(rPr, 'sz'), 1800) / 100;
  if (atr(rPr, 'b') != null) r.b = atr(rPr, 'b') === '1' || atr(rPr, 'b') === 'true';
  if (atr(rPr, 'i') != null) r.i = atr(rPr, 'i') === '1' || atr(rPr, 'i') === 'true';
  if (atr(rPr, 'u') != null) r.u = atr(rPr, 'u') !== 'none';
  if (atr(rPr, 'strike') != null) r.riscado = atr(rPr, 'strike') !== 'noStrike';
  if (atr(rPr, 'cap') != null) r.caps = atr(rPr, 'cap') === 'all';
  if (atr(rPr, 'baseline') != null) { const b = num(atr(rPr, 'baseline'), 0); r.vert = b > 0 ? 'superscript' : b < 0 ? 'subscript' : null; }
  const sf = filho(rPr, 'solidFill'); if (sf) { const c = corDrawing(sf, ctx); if (c) r.cor = c.rgb; }
  const lat = filho(rPr, 'latin');
  if (lat) { let tf = atr(lat, 'typeface'); if (tf === '+mj-lt') tf = ctx.fontes.maior; else if (tf === '+mn-lt') tf = ctx.fontes.menor; if (tf) r.fonte = tf; }
  return r;
}
function numeroAuto(n, esquema) {
  const base = /alphaLc/.test(esquema) ? letras(n).toLowerCase() : /alphaUc/.test(esquema) ? letras(n) : /romanLc/.test(esquema) ? romano(n).toLowerCase() : /romanUc/.test(esquema) ? romano(n) : String(n);
  if (/ParenBoth/.test(esquema)) return '(' + base + ')';
  if (/ParenR/.test(esquema)) return base + ')';
  if (/Period/.test(esquema)) return base + '.';
  return base;
}

class LeitorPptx {
  constructor(zip) { this.zip = zip; this.imagens = new Set(); this.aviso = { imagens: 0, objetos: 0 }; this.cacheParte = new Map(); }
  async parte(caminhoXml) {
    if (this.cacheParte.has(caminhoXml)) return this.cacheParte.get(caminhoXml);
    const d = await this.zip.xml(caminhoXml);
    const rels = await leRelacoes(this.zip, caminhoXml);
    const p = { doc: d, rels, caminho: caminhoXml };
    this.cacheParte.set(caminhoXml, p);
    return p;
  }
  async tema(caminhoTema) {
    const d = caminhoTema && await this.zip.xml(caminhoTema);
    const cores = {}, fontes = { maior: 'Calibri Light', menor: 'Calibri' };
    if (d) {
      const cs = descendentes(d.documentElement, 'clrScheme')[0];
      if (cs) for (const c of filhos(cs)) { const cc = corDrawing(c, { mapa: {}, tema: {} }); if (cc) cores[c.localName] = cc.rgb; }
      const fs = descendentes(d.documentElement, 'fontScheme')[0];
      const mj = fs && caminho(fs, 'majorFont', 'latin'), mn = fs && caminho(fs, 'minorFont', 'latin');
      if (mj) fontes.maior = atr(mj, 'typeface'); if (mn) fontes.menor = atr(mn, 'typeface');
    }
    return { cores, fontes };
  }
  placeholders(spTree) {
    const out = [];
    for (const sp of descendentes(spTree, 'sp')) {
      const ph = descendentes(filho(sp, 'nvSpPr'), 'ph')[0];
      if (ph) out.push({ sp, tipo: atr(ph, 'type') || 'body', idx: atr(ph, 'idx') });
    }
    return out;
  }
  achaPh(lista, tipo, idx) {
    if (!lista) return null;
    const norm = (t) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
    if (idx != null) { const a = lista.find((p) => p.idx === idx); if (a) return a; }
    return lista.find((p) => p.tipo === tipo) || lista.find((p) => norm(p.tipo) === norm(tipo)) || null;
  }
}

export async function lePptx(bytes) {
  const zip = await abreZip(bytes);
  const raiz = await leRelacoes(zip, '');
  let presPath = null;
  for (const [, r] of raiz) if (r.tipo === 'officeDocument') presPath = r.alvo;
  if (!presPath || !zip.tem(presPath)) throw new ErroConversao('NAO_PPTX');
  const L = new LeitorPptx(zip);
  const pres = await L.parte(presPath);
  if (!pres.doc || pres.doc.documentElement.localName !== 'presentation') throw new ErroConversao('NAO_PPTX');
  const P = pres.doc.documentElement;
  const sz = filho(P, 'sldSz');
  const W = sz ? num(atr(sz, 'cx'), 9144000) / EMU : 720, H = sz ? num(atr(sz, 'cy'), 6858000) / EMU : 540;
  const padraoTexto = leNiveis(filho(P, 'defaultTextStyle'));
  const slides = [];
  const lista = filhos(filho(P, 'sldIdLst'), 'sldId');
  for (const s of lista) {
    if (atr(s, 'show') === '0') continue;
    const rel = pres.rels.get(s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || atr(s, 'id'));
    if (!rel) continue;
    const sl = await L.parte(rel.alvo);
    if (!sl.doc) continue;
    if (atr(sl.doc.documentElement, 'show') === '0') continue;
    let lay = null, mas = null;
    for (const [, r] of sl.rels) if (r.tipo === 'slideLayout') lay = await L.parte(r.alvo);
    if (lay) for (const [, r] of lay.rels) if (r.tipo === 'slideMaster') mas = await L.parte(r.alvo);
    let tema = { cores: {}, fontes: { maior: 'Calibri', menor: 'Calibri' } };
    if (mas) for (const [, r] of mas.rels) if (r.tipo === 'theme') tema = await L.tema(r.alvo);
    const clr = mas && filho(mas.doc.documentElement, 'clrMap');
    const mapa = {};
    if (clr) for (let i = 0; i < clr.attributes.length; i++) mapa[clr.attributes[i].localName] = clr.attributes[i].value;
    const ovr = caminho(sl.doc.documentElement, 'clrMapOvr', 'overrideClrMapping');
    if (ovr) for (let i = 0; i < ovr.attributes.length; i++) mapa[ovr.attributes[i].localName] = ovr.attributes[i].value;
    const ctx = { mapa, tema: tema.cores, fontes: tema.fontes };
    const M = mas && mas.doc.documentElement, LY = lay && lay.doc.documentElement, S = sl.doc.documentElement;
    const txStyles = M && filho(M, 'txStyles');
    const estilosMestre = { title: leNiveis(txStyles && filho(txStyles, 'titleStyle')), body: leNiveis(txStyles && filho(txStyles, 'bodyStyle')), other: leNiveis(txStyles && filho(txStyles, 'otherStyle')) };
    const phMestre = M ? L.placeholders(caminho(M, 'cSld', 'spTree')) : [];
    const phLayout = LY ? L.placeholders(caminho(LY, 'cSld', 'spTree')) : [];
    /* fundo */
    let fundo = null;
    for (const fonteBg of [S, LY, M]) {
      const bg = fonteBg && caminho(fonteBg, 'cSld', 'bg');
      if (!bg) continue;
      const parteBg = fonteBg === S ? sl : fonteBg === LY ? lay : mas;
      const bgPr = filho(bg, 'bgPr'), bgRef = filho(bg, 'bgRef');
      if (bgPr) {
        const bf = filho(bgPr, 'blipFill');
        const blip = bf && filho(bf, 'blip');
        if (blip) { const r = parteBg.rels.get(atr(blip, 'embed')); if (r) { L.imagens.add(r.alvo); fundo = { img: r.alvo }; } }
        if (!fundo) { const f = preenchimento(bgPr, ctx); if (f) fundo = { cor: f.rgb }; }
      } else if (bgRef) { const c = corDrawing(bgRef, ctx); if (c) fundo = { cor: c.rgb }; }
      if (fundo) break;
    }
    if (!fundo) fundo = { cor: tema.cores[mapa.bg1 || 'lt1'] || [255, 255, 255] };
    const formas = [];
    const mostraMestre = (LY ? atr(LY, 'showMasterSp') !== '0' : true) && atr(S, 'showMasterSp') !== '0';
    const ctxForma = { L, ctx, estilosMestre, phMestre, phLayout, padraoTexto, formas };
    if (M && mostraMestre) await percorre(caminho(M, 'cSld', 'spTree'), mas, ctxForma, 'mestre', null);
    if (LY) await percorre(caminho(LY, 'cSld', 'spTree'), lay, ctxForma, 'layout', null);
    await percorre(caminho(S, 'cSld', 'spTree'), sl, ctxForma, 'slide', null);
    slides.push({ fundo, formas });
    await pausa();
  }
  const imagens = await imagensDoPacote(zip, [...L.imagens], L.aviso);
  return { tipo: 'slides', w: W, h: H, slides, imagens, avisos: L.aviso };
}

/* transforma coordenadas de filhos de um grupo */
function aplicaGrupo(g, grupo) {
  if (!grupo) return g;
  const { x, y, sx, sy } = grupo;
  return Object.assign({}, g, { x: x + g.x * sx, y: y + g.y * sy, w: g.w * sx, h: g.h * sy });
}
async function percorre(arvore, parte, c, nivel, grupo) {
  if (!arvore) return;
  for (const el of filhos(arvore)) {
    try {
      const n = el.localName;
      if (n === 'sp' || n === 'cxnSp') await forma(el, parte, c, nivel, grupo);
      else if (n === 'pic') await figura(el, parte, c, grupo);
      else if (n === 'grpSp') {
        const g = geometria(filho(el, 'grpSpPr'));
        let sub = grupo;
        if (g && g.chOff && g.chExt) {
          const cx = num(atr(g.chOff, 'x'), 0) / EMU, cy = num(atr(g.chOff, 'y'), 0) / EMU, cw = num(atr(g.chExt, 'cx'), 1) / EMU || 1, ch = num(atr(g.chExt, 'cy'), 1) / EMU || 1;
          const sx = g.w / cw, sy = g.h / ch;
          const real = aplicaGrupo(g, grupo);
          const fx = grupo ? grupo.sx : 1, fy = grupo ? grupo.sy : 1;
          sub = { x: real.x - cx * sx * fx, y: real.y - cy * sy * fy, sx: sx * fx, sy: sy * fy };
        }
        await percorre(el, parte, c, nivel, sub);
      } else if (n === 'graphicFrame') await quadro(el, parte, c, grupo);
      else if (n === 'AlternateContent') await percorre(filho(el, 'Choice') || filho(el, 'Fallback'), parte, c, nivel, grupo);
    } catch (e) { if (e instanceof ErroConversao) throw e; c.L.aviso.objetos++; }
  }
}
async function forma(el, parte, c, nivel, grupo) {
  const nv = filho(el, 'nvSpPr') || filho(el, 'nvCxnSpPr');
  const ph = nv && descendentes(nv, 'ph')[0];
  /* nos slides mestre e de layout, os espaços reservados são só moldes */
  if (ph && nivel !== 'slide') return;
  const spPr = filho(el, 'spPr');
  let g = geometria(spPr);
  let tipoPh = null, phL = null, phM = null;
  if (ph) {
    tipoPh = atr(ph, 'type') || 'body';
    const idx = atr(ph, 'idx');
    phL = c.L.achaPh(c.phLayout, tipoPh, idx);
    phM = c.L.achaPh(c.phMestre, tipoPh, idx) || (phL && c.L.achaPh(c.phMestre, phL.tipo, null));
    if (!g) g = geometria(phL && filho(phL.sp, 'spPr')) || geometria(phM && filho(phM.sp, 'spPr'));
    if (['dt', 'ftr', 'sldNum'].includes(tipoPh) && !descendentes(filho(el, 'txBody'), 't').some((t) => t.textContent.trim())) return;
  }
  if (!g) return;
  g = aplicaGrupo(g, grupo);
  const estilo = filho(el, 'style');
  const geom = filho(spPr, 'prstGeom'), cust = filho(spPr, 'custGeom');
  let fill = preenchimento(spPr, c.ctx, estilo);
  if (fill === undefined && ph) fill = preenchimento(phL && filho(phL.sp, 'spPr'), c.ctx) ?? preenchimento(phM && filho(phM.sp, 'spPr'), c.ctx) ?? null;
  if (fill === undefined) fill = null;
  let ln = contorno(spPr, c.ctx, estilo);
  if (ln === undefined) ln = null;
  const f = { tipo: 'forma', x: g.x, y: g.y, w: g.w, h: g.h, rot: g.rot, fh: g.fh, fv: g.fv, geo: geom ? atr(geom, 'prst') : cust ? 'custom' : 'rect', fill, ln };
  if (cust) f.caminho = caminhoCustom(cust, g.w, g.h);
  if (el.localName === 'cxnSp' || /line|Connector/i.test(f.geo)) { f.geo = 'linha'; if (!f.ln) f.ln = { cor: { rgb: [0, 0, 0], alfa: 1 }, larg: 0.75 }; }
  const tx = filho(el, 'txBody');
  if (tx) f.texto = texto(tx, c, tipoPh, phL, phM, el);
  if (!f.fill && !f.ln && !(f.texto && f.texto.pars.some((p) => p.runs.some((r) => r.t && r.t.trim())))) return;
  c.formas.push(f);
}
function caminhoCustom(cust, wPt, hPt) {
  const out = [];
  for (const path of descendentes(cust, 'path')) {
    const pw = num(atr(path, 'w'), 0) || wPt * EMU, phh = num(atr(path, 'h'), 0) || hPt * EMU;
    const sx = wPt / (pw / EMU) / EMU, sy = hPt / (phh / EMU) / EMU;
    let d = '';
    const pt = (e) => { const p = filho(e, 'pt') || e; return (num(atr(p, 'x'), 0) * sx).toFixed(2) + ' ' + (num(atr(p, 'y'), 0) * sy).toFixed(2); };
    for (const cmd of filhos(path)) {
      const n = cmd.localName;
      if (n === 'moveTo') d += 'M' + pt(cmd) + ' ';
      else if (n === 'lnTo') d += 'L' + pt(cmd) + ' ';
      else if (n === 'cubicBezTo') { const ps = filhos(cmd, 'pt'); if (ps.length === 3) d += 'C' + ps.map((p) => (num(atr(p, 'x'), 0) * sx).toFixed(2) + ' ' + (num(atr(p, 'y'), 0) * sy).toFixed(2)).join(' ') + ' '; }
      else if (n === 'quadBezTo') { const ps = filhos(cmd, 'pt'); if (ps.length === 2) d += 'Q' + ps.map((p) => (num(atr(p, 'x'), 0) * sx).toFixed(2) + ' ' + (num(atr(p, 'y'), 0) * sy).toFixed(2)).join(' ') + ' '; }
      else if (n === 'close') d += 'Z ';
    }
    if (d) out.push(d);
  }
  return out.join(' ');
}
function texto(tx, c, tipoPh, phL, phM, el) {
  const bp = filho(tx, 'bodyPr');
  const bpL = phL && caminho(phL.sp, 'txBody', 'bodyPr'), bpM = phM && caminho(phM.sp, 'txBody', 'bodyPr');
  const pega = (n, pad) => { for (const b of [bp, bpL, bpM]) if (b && atr(b, n) != null) return atr(b, n); return pad; };
  const ancora = pega('anchor', tipoPh === 'title' || tipoPh === 'ctrTitle' ? (tipoPh === 'ctrTitle' ? 'b' : 'ctr') : 't');
  const ins = { l: num(pega('lIns', 91440), 91440) / EMU, r: num(pega('rIns', 91440), 91440) / EMU, t: num(pega('tIns', 45720), 45720) / EMU, b: num(pega('bIns', 45720), 45720) / EMU };
  let escala = 1, reducao = 0;
  for (const b of [bp, bpL]) { const na = b && filho(b, 'normAutofit'); if (na) { if (atr(na, 'fontScale')) escala = num(atr(na, 'fontScale'), 100000) / 100000; if (atr(na, 'lnSpcReduction')) reducao = num(atr(na, 'lnSpcReduction'), 0) / 100000; break; } }
  const quebra = pega('wrap', 'square') !== 'none';
  const ehTitulo = tipoPh === 'title' || tipoPh === 'ctrTitle';
  const base = tipoPh ? (ehTitulo ? c.estilosMestre.title : ['body', 'subTitle', 'obj'].includes(tipoPh) || !tipoPh ? c.estilosMestre.body : c.estilosMestre.other) : c.estilosMestre.other;
  const camadas = [c.padraoTexto, base, phM && leNiveis(caminho(phM.sp, 'txBody', 'lstStyle')), phL && leNiveis(caminho(phL.sp, 'txBody', 'lstStyle')), leNiveis(filho(tx, 'lstStyle'))].filter(Boolean);
  /* cor do texto vinda do estilo da forma (fontRef) */
  const estilo = filho(el, 'style');
  const fontRef = estilo && filho(estilo, 'fontRef');
  const corRef = fontRef ? corDrawing(fontRef, c.ctx) : null;
  const pars = [];
  const contadores = [];
  for (const p of filhos(tx, 'p')) {
    const pPr = filho(p, 'pPr');
    const lvl = pPr ? num(atr(pPr, 'lvl'), 0) : 0;
    const props = { alin: 'left', esq: 0, primeira: 0, r: {} };
    for (const cam of camadas) { const pn = propsNivel(cam[lvl], c.ctx); const r = Object.assign({}, props.r, pn.r || {}); Object.assign(props, pn); props.r = r; }
    const direto = propsNivel(pPr, c.ctx);
    const rd = Object.assign({}, props.r, direto.r || {});
    Object.assign(props, direto); props.r = rd;
    if (corRef && !props.r.cor) props.r.cor = corRef.rgb;
    const runs = [];
    const fim = filho(p, 'endParaRPr');
    for (const r of filhos(p)) {
      const n = r.localName;
      if (n === 'r' || n === 'fld') {
        const rp = Object.assign({}, props.r, propsRun(filho(r, 'rPr'), c.ctx));
        const t = filho(r, 't');
        let s = t ? t.textContent : '';
        if (n === 'fld' && atr(r, 'type') === 'slidenum') s = '#';
        runs.push(runPdf(rp, s, escala));
      } else if (n === 'br') runs.push(Object.assign(runPdf(Object.assign({}, props.r, propsRun(filho(r, 'rPr'), c.ctx)), '', escala), { br: true }));
    }
    const vazio = !runs.some((r) => r.t);
    const rFim = runPdf(Object.assign({}, props.r, propsRun(fim, c.ctx)), '', escala);
    const tamLinha = vazio ? rFim.tam : Math.max(...runs.filter((r) => r.t).map((r) => r.tam));
    const espaco = (e, pad) => (!e ? pad : e.pts != null ? e.pts : e.pct * tamLinha * 1.2);
    let marcador = null;
    if (!vazio && props.bullet) {
      if (props.bullet.tipo === 'char') marcador = { texto: /[-]/.test(props.bullet.c) || props.bullet.c.length !== 1 ? '•' : props.bullet.c, r: Object.assign({}, runs[0], { cor: props.buCor ? props.buCor.rgb : runs[0].cor, b: false, i: false, u: false }) };
      else { contadores[lvl] = (contadores[lvl] || props.bullet.inicio - 1) + 1; contadores.length = lvl + 1; marcador = { texto: numeroAuto(contadores[lvl], props.bullet.esquema), r: Object.assign({}, runs[0], { u: false }) }; }
    } else if (!props.bullet) contadores.length = 0;
    const ls = props.linha;
    pars.push({
      tipo: 'p', alin: props.alin, esq: props.esq || 0, dir: 0, primeira: props.primeira || 0,
      antes: espaco(props.antes, 0), depois: espaco(props.depois, 0),
      linha: ls ? (ls.pts != null ? { tipo: 'exato', v: ls.pts } : { tipo: 'mult', v: Math.max(0.5, ls.pct * (1 - reducao)) }) : { tipo: 'mult', v: 1 - reducao },
      runs: vazio ? [] : runs, marcador, tamPadrao: rFim.tam,
    });
  }
  return { pars, ancora, ins, quebra };
}
function runPdf(r, t, escala) {
  return { t: r.caps ? t.toUpperCase() : t, b: !!r.b, i: !!r.i, u: !!r.u, riscado: !!r.riscado, tam: (r.tam || 18) * escala, cor: r.cor || [0, 0, 0], fonte: familiaPdf(r.fonte) || 'helv', vert: r.vert || null };
}
async function figura(el, parte, c, grupo) {
  let g = geometria(filho(el, 'spPr'));
  if (!g) return;
  g = aplicaGrupo(g, grupo);
  const bf = filho(el, 'blipFill');
  const blip = bf && filho(bf, 'blip');
  const rel = blip && parte.rels.get(atr(blip, 'embed'));
  if (!rel || rel.externo) { c.L.aviso.imagens++; return; }
  c.L.imagens.add(rel.alvo);
  const sr = bf && filho(bf, 'srcRect');
  const corte = sr ? { l: num(atr(sr, 'l'), 0) / 100000, t: num(atr(sr, 't'), 0) / 100000, r: num(atr(sr, 'r'), 0) / 100000, b: num(atr(sr, 'b'), 0) / 100000 } : null;
  const geom = filho(filho(el, 'spPr'), 'prstGeom');
  c.formas.push({ tipo: 'imagem', x: g.x, y: g.y, w: g.w, h: g.h, rot: g.rot, fh: g.fh, fv: g.fv, img: rel.alvo, corte, geo: geom ? atr(geom, 'prst') : 'rect', ln: contorno(filho(el, 'spPr'), c.ctx, filho(el, 'style')) || null });
}
async function quadro(el, parte, c, grupo) {
  const x = filho(el, 'xfrm');
  if (!x) return;
  let g = { x: num(atr(filho(x, 'off'), 'x'), 0) / EMU, y: num(atr(filho(x, 'off'), 'y'), 0) / EMU, w: num(atr(filho(x, 'ext'), 'cx'), 0) / EMU, h: num(atr(filho(x, 'ext'), 'cy'), 0) / EMU };
  g = aplicaGrupo(g, grupo);
  const tbl = descendentes(el, 'tbl')[0];
  if (!tbl) {
    /* gráficos e SmartArt: usa a imagem de reserva quando existir */
    const fb = descendentes(el, 'blip')[0];
    const rel = fb && parte.rels.get(atr(fb, 'embed'));
    if (rel) { c.L.imagens.add(rel.alvo); c.formas.push({ tipo: 'imagem', x: g.x, y: g.y, w: g.w, h: g.h, img: rel.alvo }); }
    else c.L.aviso.objetos++;
    return;
  }
  const cols = filhos(filho(tbl, 'tblGrid'), 'gridCol').map((gc) => num(atr(gc, 'w'), 0) / EMU);
  const linhas = [];
  const tp = filho(tbl, 'tblPr');
  const primeiraLinha = tp && atr(tp, 'firstRow') === '1';
  const faixa = tp && atr(tp, 'bandRow') === '1';
  const destaque = c.ctx.tema[c.ctx.mapa.accent1 || 'accent1'] || [68, 114, 196];
  filhos(tbl, 'tr').forEach((tr, ri) => {
    const linha = { alt: num(atr(tr, 'h'), 0) / EMU, celulas: [] };
    for (const tc of filhos(tr, 'tc')) {
      const tcPr = filho(tc, 'tcPr');
      let fundo = tcPr ? preenchimento(tcPr, c.ctx) : undefined;
      let corTexto = null;
      if (fundo === undefined) {
        if (primeiraLinha && ri === 0) { fundo = { rgb: destaque }; corTexto = [255, 255, 255]; }
        else if (faixa) fundo = { rgb: ri % 2 ? destaque.map((v) => Math.round(255 - (255 - v) * 0.2)) : destaque.map((v) => Math.round(255 - (255 - v) * 0.35)) };
        else fundo = null;
      }
      const txb = filho(tc, 'txBody');
      const tt = txb ? texto(txb, c, null, null, null, tc) : { pars: [], ancora: 't', ins: { l: 7.2, r: 7.2, t: 3.6, b: 3.6 } };
      if (corTexto) for (const p of tt.pars) for (const r of p.runs) if (!r.corDefinida) r.cor = corTexto;
      if (tcPr) { tt.ins = { l: num(atr(tcPr, 'marL'), 91440) / EMU, r: num(atr(tcPr, 'marR'), 91440) / EMU, t: num(atr(tcPr, 'marT'), 45720) / EMU, b: num(atr(tcPr, 'marB'), 45720) / EMU }; const an = atr(tcPr, 'anchor'); tt.ancora = an || 't'; }
      linha.celulas.push({ span: num(atr(tc, 'gridSpan'), 1), rspan: num(atr(tc, 'rowSpan'), 1), oculta: atr(tc, 'hMerge') === '1' || atr(tc, 'vMerge') === '1', fundo: fundo ? fundo.rgb : null, texto: tt });
    }
    linhas.push(linha);
  });
  c.formas.push({ tipo: 'tabela', x: g.x, y: g.y, w: g.w, h: g.h, cols, linhas });
}
