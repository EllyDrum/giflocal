/* GIF Local, conversões para PDF: diagramação fora da tela.
   Recebe o "modelo" lido de um .docx, .xlsx/.csv ou .pptx (feito em
   conv-para-pdf.js) e desenha as páginas com a pdf-lib (MIT).
   Nenhuma chamada de rede: entra o modelo, sai o PDF.
   As fontes são as 14 fontes padrão do PDF (Helvetica, Times, Courier):
   não é preciso embutir arquivos de fonte, e o PDF fica leve. Letras fora
   do conjunto latino (ex.: grego, cirílico, emojis) não existem nessas
   fontes e são trocadas por equivalentes ou por "?"; a interface avisa. */
'use strict';
importScripts('/pdf/vendor/pdf-lib-2.11.1/pdf-lib.min.js');
const { PDFDocument, StandardFonts, rgb, degrees, pushGraphicsState, popGraphicsState, concatTransformationMatrix, rectangle, clip, endPath } = self.PDFLib;
const PRODUTOR = 'GIF Local (giflocal.com)';

class Falha extends Error { constructor(codigo, detalhe) { super(codigo); this.codigo = codigo; this.detalhe = detalhe || null; } }
function progresso(id, feito, total) { self.postMessage({ id, tipo: 'progresso', feito, total }); }

/* ---------------- fontes e caracteres ---------------- */
let F = null, CONJUNTO = null, trocados = 0;
const SUBST = { '→': '->', '←': '<-', '↔': '<->', '⇒': '=>', '≤': '<=', '≥': '>=', '≠': '!=', '−': '-', '‐': '-', '‑': '-', '‒': '-', '―': '—', '′': "'", '″': '"', '✓': 'v', '✔': 'v', '✗': 'x', '✘': 'x', '●': '•', '○': 'o', '▪': '•', '■': '•', '□': 'o', '◦': 'o', '‣': '•', '⁃': '-', '∙': '•', '⋅': '·', '​': '', '‌': '', '‍': '', '﻿': '', ' ': ' ', ' ': ' ', '­': '', ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', '∶': ':', '∕': '/', ' ': ' ', '№': 'No', '™': '™', '…': '…', '≈': '~', '×': '×', '∗': '*', '★': '*', '☆': '*', '✔️': 'v' };
function limpa(t) {
  let out = '';
  for (const ch of t) {
    const cp = ch.codePointAt(0);
    if (CONJUNTO.has(cp)) { out += ch; continue; }
    if (SUBST[ch] !== undefined) { out += SUBST[ch]; continue; }
    if (cp === 9) { out += ' '; continue; }
    if (cp < 32 || (cp >= 0xFE00 && cp <= 0xFE0F)) continue;
    const base = ch.normalize('NFD')[0];
    if (base && CONJUNTO.has(base.codePointAt(0)) && base !== ch) { out += base; continue; }
    trocados++;
    out += '?';
  }
  return out;
}
async function carregaFontes(doc) {
  const n = StandardFonts;
  const tab = {
    helv: [n.Helvetica, n.HelveticaBold, n.HelveticaOblique, n.HelveticaBoldOblique],
    times: [n.TimesRoman, n.TimesRomanBold, n.TimesRomanItalic, n.TimesRomanBoldItalic],
    courier: [n.Courier, n.CourierBold, n.CourierOblique, n.CourierBoldOblique],
  };
  F = {};
  for (const [k, lista] of Object.entries(tab)) F[k] = await Promise.all(lista.map((x) => doc.embedFont(x)));
  CONJUNTO = new Set(F.helv[0].getCharacterSet());
}
function fonte(r) { const f = F[r.fonte] || F.helv; return f[(r.b ? 1 : 0) + (r.i ? 2 : 0)]; }
function cor(c) { return c ? rgb(Math.min(255, c[0]) / 255, Math.min(255, c[1]) / 255, Math.min(255, c[2]) / 255) : rgb(0, 0, 0); }
function tamReal(r) { return r.vert === 'superscript' || r.vert === 'subscript' ? r.tam * 0.65 : r.tam; }
function largura(r, t) { if (!t) return 0; try { return fonte(r).widthOfTextAtSize(t, tamReal(r)); } catch (e) { return t.length * tamReal(r) * 0.5; } }

/* ---------------- imagens ---------------- */
let IMGS = {}, cacheImg = new Map();
async function imagem(doc, id) {
  if (cacheImg.has(id)) return cacheImg.get(id);
  const im = IMGS[id];
  let r = null;
  if (im) { try { r = im.tipo === 'jpg' ? await doc.embedJpg(im.bytes) : await doc.embedPng(im.bytes); } catch (e) { r = null; } }
  cacheImg.set(id, r);
  return r;
}

/* ================================================================
   Parágrafo: quebra em linhas
   ================================================================ */
/* p: { runs, alin, esq, dir, primeira, linha, marcador, tamPadrao }
   devolve [{ itens:[{tipo, t, r, w, img}], w, h, base, ultima, recuo }] */
function quebraParagrafo(p, larg, fator) {
  const tokens = [];
  for (const r of p.runs) {
    if (r.quebraPagina) { tokens.push({ tipo: 'quebraPagina' }); continue; }
    if (r.br) { tokens.push({ tipo: 'br', r }); continue; }
    if (r.tab) { tokens.push({ tipo: 'tab', r }); continue; }
    if (r.img) { tokens.push({ tipo: 'img', r, img: r.img }); continue; }
    if (r.campo) { tokens.push({ tipo: 'texto', r, t: '00', campo: r.campo }); continue; }
    const t = limpa(r.t || '');
    if (!t) continue;
    for (const parte of t.split(/( +)/)) if (parte) tokens.push({ tipo: parte[0] === ' ' ? 'espaco' : 'texto', r, t: parte });
  }
  const recuoMarcador = p.marcador ? p.esq + p.primeira : null;
  let larguraMarcador = 0;
  let inicioPrimeira = p.esq + p.primeira;
  if (p.marcador) {
    larguraMarcador = largura(p.marcador.r, limpa(p.marcador.texto));
    inicioPrimeira = p.primeira < 0 && recuoMarcador + larguraMarcador + 3 <= p.esq ? p.esq : recuoMarcador + larguraMarcador + Math.max(4, p.marcador.r.tam * 0.4);
  }
  const linhas = [];
  let linha = null;
  const nova = () => { linha = { itens: [], w: 0, recuo: linhas.length === 0 ? inicioPrimeira : p.esq }; linha.disp = Math.max(10, larg - linha.recuo - (p.dir || 0)); linhas.push(linha); };
  nova();
  const fechaEspacos = (L) => { while (L.itens.length && L.itens[L.itens.length - 1].tipo === 'espaco') L.w -= L.itens.pop().w; };
  for (const tk of tokens) {
    if (tk.tipo === 'quebraPagina') { fechaEspacos(linha); linha.quebraDepois = true; nova(); continue; }
    if (tk.tipo === 'br') { fechaEspacos(linha); linha.forcada = true; linha.rBr = tk.r; nova(); continue; }
    if (tk.tipo === 'espaco') {
      if (!linha.itens.length && linhas.length > 1) continue;
      const w = largura(tk.r, tk.t);
      linha.itens.push({ tipo: 'espaco', t: tk.t, r: tk.r, w });
      linha.w += w;
      continue;
    }
    if (tk.tipo === 'tab') {
      const pos = linha.recuo + linha.w;
      const alvo = (Math.floor(pos / 36) + 1) * 36;
      const w = Math.max(4, alvo - pos);
      if (linha.w + w > linha.disp && linha.itens.length) { fechaEspacos(linha); nova(); }
      linha.itens.push({ tipo: 'espaco', t: ' ', r: tk.r, w, tab: true });
      linha.w += w;
      continue;
    }
    if (tk.tipo === 'img') {
      let w = tk.img.w, h = tk.img.h;
      const k = Math.min(1, (larg - p.esq) / w, (fator.altMax || 1e9) / h);
      w *= k; h *= k;
      if (linha.w + w > linha.disp && linha.itens.length) { fechaEspacos(linha); nova(); }
      linha.itens.push({ tipo: 'img', img: tk.img, w, h, r: tk.r });
      linha.w += w;
      continue;
    }
    let w = largura(tk.r, tk.t);
    if (linha.w + w > linha.disp && linha.itens.some((i) => i.tipo !== 'espaco')) {
      /* volta as letras grudadas no fim da linha (ex.: "palavra," vinda de outro run) */
      let k = linha.itens.length;
      while (k > 0 && linha.itens[k - 1].tipo === 'texto') k--;
      const levar = k > 0 ? linha.itens.splice(k) : [];
      linha.w -= levar.reduce((a, x) => a + x.w, 0);
      fechaEspacos(linha);
      nova();
      for (const x of levar) { linha.itens.push(x); linha.w += x.w; }
    }
    if (w > linha.disp && !linha.itens.length) {
      /* palavra maior que a linha: corta por letras */
      let resto = tk.t;
      while (resto) {
        let n = resto.length;
        while (n > 1 && largura(tk.r, resto.slice(0, n)) > linha.disp - linha.w) n--;
        const pedaco = resto.slice(0, n);
        const wp = largura(tk.r, pedaco);
        linha.itens.push({ tipo: 'texto', t: pedaco, r: tk.r, w: wp, campo: tk.campo });
        linha.w += wp;
        resto = resto.slice(n);
        if (resto) nova();
      }
      continue;
    }
    linha.itens.push({ tipo: 'texto', t: tk.t, r: tk.r, w, campo: tk.campo });
    linha.w += w;
  }
  fechaEspacos(linha);
  if (linhas.length > 1 && !linha.itens.length && !linhas[linhas.length - 2].forcada && !linhas[linhas.length - 2].quebraDepois) linhas.pop();
  linhas[linhas.length - 1].ultima = true;
  /* alturas */
  for (const L of linhas) {
    let tam = 0, img = 0;
    for (const it of L.itens) { if (it.tipo === 'img') img = Math.max(img, it.h); else tam = Math.max(tam, it.r.tam); }
    if (!tam) tam = (L.rBr && L.rBr.tam) || p.tamPadrao || 11;
    const natural = tam * (fator.natural || 1.17);
    const lin = p.linha || { tipo: 'mult', v: 1 };
    let h = lin.tipo === 'exato' ? lin.v : lin.tipo === 'min' ? Math.max(lin.v, natural) : natural * lin.v;
    /* linha de base: o espaço extra (entrelinha maior) fica metade acima, metade abaixo */
    const extra = h - natural;
    L.base = natural - tam * 0.27 + (extra > 0 ? extra * 0.5 : extra);
    if (img) { h = Math.max(h, img + tam * 0.3); L.base = img; }
    if (L.quebraDepois && !L.itens.length) h = 0; // a quebra de página em si não ocupa espaço
    L.h = h;
    L.tam = tam;
  }
  if (p.marcador) linhas[0].marcador = { t: limpa(p.marcador.texto), r: p.marcador.r, x: recuoMarcador };
  return linhas;
}

/* desenha uma linha; x0 = início da área do parágrafo; yTopo em coordenadas de cima para baixo */
async function desenhaLinha(ctx, L, p, x0, larg, yTopo, campos) {
  const { page, doc, H } = ctx;
  const yBase = H - (yTopo + L.base);
  let x = x0 + L.recuo;
  const sobra = L.disp - L.w;
  let extraEspaco = 0;
  if (p.alin === 'center') x += sobra / 2;
  else if (p.alin === 'right') x += sobra;
  else if (p.alin === 'both' && !L.ultima && !L.forcada) {
    const n = L.itens.filter((i) => i.tipo === 'espaco' && !i.tab).length;
    if (n && sobra > 0 && sobra < L.disp * 0.5) extraEspaco = sobra / n;
  }
  if (L.marcador) {
    const m = L.marcador;
    page.drawText(m.t, { x: x0 + m.x, y: yBase, size: tamReal(m.r), font: fonte(m.r), color: cor(m.r.cor) });
  }
  for (const it of L.itens) {
    if (it.tipo === 'espaco') {
      const w = it.w + (it.tab ? 0 : extraEspaco);
      if (it.r.u && !it.tab) page.drawLine({ start: { x, y: yBase - it.r.tam * 0.12 }, end: { x: x + w, y: yBase - it.r.tam * 0.12 }, thickness: Math.max(0.4, it.r.tam * 0.05), color: cor(it.r.cor) });
      x += w;
      continue;
    }
    if (it.tipo === 'img') {
      const im = await imagem(doc, it.img.id);
      if (im) page.drawImage(im, { x, y: yBase, width: it.w, height: it.h });
      x += it.w;
      continue;
    }
    const r = it.r;
    let t = it.t;
    if (it.campo) { if (campos) t = String(it.campo === 'PAGE' ? campos.pagina : campos.total); else { ctx.camposPendentes = true; } }
    const tam = tamReal(r);
    const dy = r.vert === 'superscript' ? r.tam * 0.33 : r.vert === 'subscript' ? -r.tam * 0.12 : 0;
    const w = it.campo ? largura(r, t) : it.w;
    if (r.realce) page.drawRectangle({ x, y: yBase - r.tam * 0.22, width: w, height: r.tam * 1.05, color: cor(r.realce) });
    page.drawText(t, { x, y: yBase + dy, size: tam, font: fonte(r), color: cor(r.cor) });
    if (r.u) page.drawLine({ start: { x, y: yBase - r.tam * 0.12 }, end: { x: x + w, y: yBase - r.tam * 0.12 }, thickness: Math.max(0.4, r.tam * 0.05), color: cor(r.cor) });
    if (r.riscado) page.drawLine({ start: { x, y: yBase + r.tam * 0.3 }, end: { x: x + w, y: yBase + r.tam * 0.3 }, thickness: Math.max(0.4, r.tam * 0.05), color: cor(r.cor) });
    x += w;
  }
}

/* ================================================================
   DOCUMENTO (Word)
   ================================================================ */
/* Transforma blocos em "itens" de altura conhecida, prontos para paginar. */
function itensDosBlocos(blocos, larg, opc) {
  const itens = [];
  blocos.forEach((b, bi) => {
    if (b.tipo === 'p') {
      if (b.quebraAntes) itens.push({ tipo: 'quebra' });
      /* espaçamento contextual: parágrafos seguidos do mesmo estilo (ex.: listas) ficam juntos */
      const ant = blocos[bi - 1], prox = blocos[bi + 1];
      const juntaAntes = b.contextual && ant && ant.tipo === 'p' && ant.estiloId === b.estiloId;
      const juntaDepois = b.contextual && prox && prox.tipo === 'p' && prox.estiloId === b.estiloId;
      if (b.antes && !juntaAntes) itens.push({ tipo: 'espaco', h: b.antes });
      const linhas = quebraParagrafo(b, larg, opc);
      if (b.fundo) itens.push({ tipo: 'fundo', cor: b.fundo, h: 0, alvo: linhas.reduce((a, L) => a + L.h, 0) });
      linhas.forEach((L, i) => {
        itens.push({ tipo: 'linha', h: L.h, L, p: b, manter: b.manterProximo || (i === 0 && linhas.length > 1 && linhas.length <= 3) });
        if (L.quebraDepois) itens.push({ tipo: 'quebra' });
      });
      if (b.depois && !juntaDepois) itens.push({ tipo: 'espaco', h: b.depois });
    } else if (b.tipo === 'tabela') {
      for (const it of itensTabela(b, larg, opc)) itens.push(it);
      if (opc.aposTabela !== false) itens.push({ tipo: 'espaco', h: 4 });
    }
  });
  return itens;
}
function largurasTabela(t, larg) {
  let cols = t.cols.slice();
  const soma = cols.reduce((a, c) => a + c, 0);
  if (!soma) cols = cols.map(() => larg / cols.length);
  else if (soma > larg) cols = cols.map((c) => c * larg / soma);
  return cols;
}
function itensTabela(t, larg, opc) {
  const cols = largurasTabela(t, larg - (t.recuo > 0 && t.recuo < larg / 2 ? t.recuo : 0));
  const total = cols.reduce((a, c) => a + c, 0);
  const desloc = t.alin === 'center' ? (larg - total) / 2 : t.alin === 'right' ? larg - total : (t.recuo > 0 && t.recuo < larg / 2 ? t.recuo : 0);
  const pad = t.margem != null ? t.margem : 5.4;
  const linhas = [];
  for (const lin of t.linhas) {
    let c = 0;
    const cels = [];
    for (const cel of lin.celulas) {
      const span = Math.max(1, cel.span || 1);
      const w = cols.slice(c, c + span).reduce((a, x) => a + x, 0) || cols[cols.length - 1] || 50;
      const x = cols.slice(0, c).reduce((a, x2) => a + x2, 0);
      const conteudo = cel.vmerge === 'continua' || cel.vazia ? [] : itensDosBlocos(cel.blocos, Math.max(10, w - 2 * pad), Object.assign({}, opc, { aposTabela: false }))
        .filter((i) => i.tipo !== 'quebra');
      while (conteudo.length && conteudo[conteudo.length - 1].tipo === 'espaco') conteudo.pop();
      const h = conteudo.reduce((a, i) => a + (i.h || 0), 0);
      cels.push({ x, w, conteudo, h, fundo: cel.fundo, vmerge: cel.vmerge, valign: cel.valign, vazia: cel.vazia });
      c += span;
    }
    const alt = Math.max(lin.altExata ? lin.alt : 0, Math.max(lin.alt || 0, ...cels.map((x) => x.h + 2)), 6);
    linhas.push({ tipo: 'linhaTabela', h: lin.altExata ? lin.alt : alt, cels, cabecalho: lin.cabecalho, bordas: t.bordas, x: desloc, pad, tabela: t });
  }
  return linhas;
}

class Paginador {
  constructor(doc, sec, lista) { this.doc = doc; this.sec = sec; this.paginas = lista || []; this.page = null; this.y = 0; this.cabecalhos = []; this.paginaDaSecao = 0; }
  novaPagina() {
    const s = this.sec;
    this.page = this.doc.addPage([s.w, s.h]);
    this.paginas.push({ page: this.page, sec: s, primeira: this.paginaDaSecao === 0, modelo: this.modelo });
    this.paginaDaSecao++;
    this.y = s.t;
    this.noTopo = true;
    for (const c of this.cabecalhos) this.desenhaLinhaTabelaInteira(c);
  }
  get ctx() { return { page: this.page, doc: this.doc, H: this.sec.h }; }
  get fundo() { return this.sec.h - this.sec.b; }
  async paginar(itens, x0, larg) {
    for (let i = 0; i < itens.length; i++) {
      const it = itens[i];
      if (it.tipo === 'quebra') { if (!this.noTopo) this.novaPagina(); continue; }
      if (it.tipo === 'espaco') { if (this.noTopo) continue; this.y += it.h; if (this.y > this.fundo) this.novaPagina(); continue; }
      if (it.tipo === 'fundo') { this.fundoPendente = it; continue; }
      if (it.tipo === 'linha') {
        let h = it.h;
        /* título não fica sozinho no fim da página */
        if (it.manter && !this.noTopo) { let k = i + 1, extra = 0; while (k < itens.length && itens[k].tipo === 'espaco') extra += itens[k++].h; if (k < itens.length && itens[k].h) extra += Math.min(itens[k].h, 60); if (this.y + h + extra > this.fundo) this.novaPagina(); }
        if (this.y + h > this.fundo + 0.5 && !this.noTopo) this.novaPagina();
        if (this.fundoPendente) { const f = this.fundoPendente; this.page.drawRectangle({ x: x0, y: this.sec.h - this.y - Math.min(f.alvo, this.fundo - this.y), width: larg, height: Math.min(f.alvo, this.fundo - this.y), color: cor(f.cor) }); this.fundoPendente = null; }
        await desenhaLinha(this.ctx, it.L, it.p, x0, larg, this.y, null);
        if (this.ctx.camposPendentes) this.temCampos = true;
        this.y += h;
        this.noTopo = false;
        continue;
      }
      if (it.tipo === 'linhaTabela') {
        if (it.cabecalho && (i === 0 || itens[i - 1].tipo !== 'linhaTabela' || itens[i - 1].cabecalho)) {
          if (!itens[i - 1] || itens[i - 1].tipo !== 'linhaTabela') this.cabecalhos = [];
          this.cabecalhos.push(it);
        } else if (!it.cabecalho && itens[i - 1] && itens[i - 1].tipo !== 'linhaTabela') this.cabecalhos = [];
        const disponivel = this.fundo - this.y;
        const cheia = this.fundo - this.sec.t;
        if (it.h <= disponivel + 0.5) { await this.desenhaLinhaTabelaInteira(it, x0); this.noTopo = false; }
        else if (it.h <= cheia * 0.9 || disponivel < cheia * 0.25) {
          if (!this.noTopo) { const cab = this.cabecalhos.includes(it) ? [] : this.cabecalhos; this.cabecalhos = cab; this.x0Tabela = x0; this.novaPagina(); }
          if (it.h <= this.fundo - this.y + 0.5) { await this.desenhaLinhaTabelaInteira(it, x0); this.noTopo = false; }
          else await this.divideLinha(it, x0);
        } else await this.divideLinha(it, x0);
        if (itens[i + 1] && itens[i + 1].tipo !== 'linhaTabela') this.cabecalhos = [];
        continue;
      }
    }
  }
  async desenhaLinhaTabelaInteira(it, x0) {
    if (x0 == null) x0 = this.x0Tabela || this.sec.l;
    this.x0Tabela = x0;
    const H = this.sec.h, yTopo = this.y, h = it.h;
    for (const c of it.cels) {
      const x = x0 + it.x + c.x;
      if (c.fundo) this.page.drawRectangle({ x, y: H - yTopo - h, width: c.w, height: h, color: cor(c.fundo) });
      if (it.bordas && !c.vazia) this.page.drawRectangle({ x, y: H - yTopo - h, width: c.w, height: h, borderColor: rgb(0.25, 0.25, 0.25), borderWidth: 0.5 });
      let y = yTopo + (c.valign === 'center' ? (h - c.h) / 2 : c.valign === 'bottom' ? h - c.h - 1 : 1);
      for (const i of c.conteudo) {
        if (i.tipo === 'linha') await desenhaLinha({ page: this.page, doc: this.doc, H }, i.L, i.p, x + it.pad, c.w - 2 * it.pad, y, null);
        else if (i.tipo === 'linhaTabela') { const yy = this.y; this.y = y; await this.desenhaLinhaTabelaInteira(i, x + it.pad); this.y = yy; this.x0Tabela = x0; }
        y += i.h || 0;
      }
    }
    this.y += h;
  }
  /* linha de tabela maior que a página: o conteúdo de cada célula continua na próxima */
  async divideLinha(it, x0) {
    const filas = it.cels.map((c) => ({ c, k: 0 }));
    const H = this.sec.h;
    let seguranca = 0;
    while (filas.some((f) => f.k < f.c.conteudo.length) && seguranca++ < 500) {
      const disponivel = this.fundo - this.y;
      let usado = 0;
      const desenhos = [];
      for (const f of filas) {
        let acum = 0;
        const lote = [];
        while (f.k < f.c.conteudo.length) {
          const i = f.c.conteudo[f.k];
          if (acum + (i.h || 0) > disponivel - 2 && lote.length) break;
          lote.push(i); acum += i.h || 0; f.k++;
          if (acum > disponivel) break;
        }
        usado = Math.max(usado, acum);
        desenhos.push({ f, lote });
      }
      const h = Math.min(disponivel, usado + 2);
      for (const { f, lote } of desenhos) {
        const c = f.c, x = x0 + it.x + c.x;
        if (c.fundo) this.page.drawRectangle({ x, y: H - this.y - h, width: c.w, height: h, color: cor(c.fundo) });
        if (it.bordas) this.page.drawRectangle({ x, y: H - this.y - h, width: c.w, height: h, borderColor: rgb(0.25, 0.25, 0.25), borderWidth: 0.5 });
        let y = this.y + 1;
        for (const i of lote) { if (i.tipo === 'linha') await desenhaLinha({ page: this.page, doc: this.doc, H }, i.L, i.p, x + it.pad, c.w - 2 * it.pad, y, null); y += i.h || 0; }
      }
      this.y += h;
      this.noTopo = false;
      if (filas.some((f) => f.k < f.c.conteudo.length)) { const cab = this.cabecalhos; this.cabecalhos = []; this.novaPagina(); this.cabecalhos = cab; }
    }
  }
}

async function documentoParaPdf(modelo, id) {
  const doc = await PDFDocument.create();
  await carregaFontes(doc);
  const lista = [];
  const pag = new Paginador(doc, null, lista);
  for (let si = 0; si < modelo.secoes.length; si++) {
    const s = modelo.secoes[si];
    const sec = s.sect;
    if (sec.l + sec.r > sec.w - 72) { sec.l = sec.r = 36; }
    if (sec.t + sec.b > sec.h - 72) { sec.t = sec.b = 36; }
    const larg = sec.w - sec.l - sec.r;
    const opc = { altMax: sec.h - sec.t - sec.b - 2 };
    const continua = si > 0 && sec.tipo === 'continuous' && pag.sec && pag.sec.w === sec.w && pag.sec.h === sec.h;
    pag.sec = sec;
    pag.modelo = s;
    if (!continua) { pag.paginaDaSecao = 0; pag.cabecalhos = []; pag.novaPagina(); }
    const itens = itensDosBlocos(s.blocos, larg, opc);
    await pag.paginar(itens, sec.l, larg);
    progresso(id, si + 1, modelo.secoes.length);
  }
  /* cabeçalhos e rodapés, depois de saber o total de páginas */
  const total = lista.length;
  for (let k = 0; k < lista.length; k++) {
    const p = lista[k], s = p.modelo, sec = p.sec;
    const larg = sec.w - sec.l - sec.r;
    const cab = p.primeira && s.sect.tituloPg ? s.cab_first : s.cab_default;
    const rod = p.primeira && s.sect.tituloPg ? s.rod_first : s.rod_default;
    const campos = { pagina: k + 1, total };
    for (const [blocos, topo] of [[cab, true], [rod, false]]) {
      if (!blocos || !blocos.length) continue;
      const itens = itensDosBlocos(blocos, larg, { altMax: sec.h / 3 }).filter((i) => i.tipo === 'linha' || i.tipo === 'espaco' || i.tipo === 'linhaTabela');
      const alt = itens.reduce((a, i) => a + (i.h || 0), 0);
      let y = topo ? sec.cab : sec.h - sec.rod - alt;
      for (const i of itens) {
        if (i.tipo === 'linha') await desenhaLinha({ page: p.page, doc, H: sec.h }, i.L, i.p, sec.l, larg, y, campos);
        else if (i.tipo === 'linhaTabela') { const tmp = new Paginador(doc, sec, []); tmp.page = p.page; tmp.y = y; await tmp.desenhaLinhaTabelaInteira(i, sec.l); }
        y += i.h || 0;
      }
    }
  }
  return doc;
}

/* ================================================================
   PLANILHAS (Excel)
   ================================================================ */
const larguraColuna = (w) => Math.max(0, Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7)) * 0.75;
const PAPEIS = { a4: [595.28, 841.89], carta: [612, 792] };

async function planilhasParaPdf(modelo, op, id) {
  const doc = await PDFDocument.create();
  await carregaFontes(doc);
  const [pw, ph] = PAPEIS[op.papel] || PAPEIS.a4;
  const margem = 28;
  const folhas = modelo.folhas.filter((f, i) => (op.folhas === 'todas' ? !f.oculta : i === op.indice) && !f.vazia && f.maxC >= 0);
  let feitas = 0;
  for (const f of folhas) {
    const cols = f.colunas.map((c) => (c.oculta ? 0 : larguraColuna(c.w)));
    const totalW = cols.reduce((a, c) => a + c, 0) || 100;
    let paisagem;
    if (op.orientacao === 'paisagem') paisagem = true;
    else if (op.orientacao === 'retrato') paisagem = false;
    else paisagem = f.orientacao === 'landscape' || totalW > (pw - 2 * margem) * 1.05;
    const W = paisagem ? ph : pw, H = paisagem ? pw : ph;
    const areaW = W - 2 * margem, areaH = H - 2 * margem - (folhas.length > 1 ? 14 : 0);
    const escala = op.ajuste === 'largura' ? Math.min(1, areaW / totalW) : 1;
    /* grupos de colunas que cabem na largura (quando não ajusta à página) */
    const grupos = [];
    let ini = 0;
    while (ini < cols.length) {
      let fim = ini, acc = 0;
      while (fim < cols.length && (acc + cols[fim] * escala <= areaW + 0.5 || fim === ini)) { acc += cols[fim] * escala; fim++; }
      grupos.push([ini, fim]);
      ini = fim;
    }
    /* mapa de células e mesclas */
    const mapa = new Map();
    for (let r = 0; r < f.linhas.length; r++) { const L = f.linhas[r]; if (L) for (const c of L.celulas) mapa.set(r * 20000 + c.c, c); }
    const cobertas = new Set(), mesclaDe = new Map();
    for (const m of f.mescla) {
      mesclaDe.set(m.r0 * 20000 + m.c0, m);
      for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) cobertas.add(r * 20000 + c);
    }
    const alturas = [];
    for (let r = 0; r <= f.maxR; r++) {
      const L = f.linhas[r];
      if (L && L.oculta) { alturas[r] = 0; continue; }
      let h = (L && L.h != null) ? L.h : f.alturaPadrao;
      if (L) for (const c of L.celulas) {
        if (c.tam && c.tam * 1.3 > h && (!L.h)) h = c.tam * 1.3;
        if (c.quebra && c.texto && !cobertas.has(r * 20000 + c.c)) {
          const linhas = quebraTexto(c, Math.max(8, cols[c.c] - 4));
          h = Math.max(h, linhas.length * c.tam * 1.2 + 3);
        }
      }
      alturas[r] = h;
    }
    const repete = op.repetir && f.maxR > 0 ? 0 : -1;
    for (const [c0, c1] of grupos) {
      let page = null, y = 0;
      const novaPagina = () => {
        page = doc.addPage([W, H]);
        y = margem;
        if (folhas.length > 1) { page.drawText(limpa(f.nome), { x: margem, y: H - margem + 6, size: 8, font: F.helv[0], color: rgb(0.45, 0.45, 0.45) }); y += 14; }
      };
      novaPagina();
      let primeira = true;
      for (let r = 0; r <= f.maxR; r++) {
        const h = alturas[r] * escala;
        if (!h) continue;
        if (y + h > H - margem && !primeira) {
          novaPagina();
          if (repete >= 0 && r > repete) { desenhaLinhaPlanilha(page, f, repete, mapa, cobertas, mesclaDe, cols, alturas, c0, c1, margem, y, H, escala, op); y += alturas[repete] * escala; }
        }
        desenhaLinhaPlanilha(page, f, r, mapa, cobertas, mesclaDe, cols, alturas, c0, c1, margem, y, H, escala, op);
        y += h;
        primeira = false;
      }
    }
    feitas++;
    progresso(id, feitas, folhas.length);
  }
  if (!doc.getPageCount()) throw new Falha('PLANILHA_VAZIA');
  return doc;
}
function quebraTexto(c, w) {
  const out = [];
  for (const par of String(c.texto).split(/\r?\n/)) {
    let linha = '';
    for (const pal of limpa(par).split(/(\s+)/)) {
      const tent = linha + pal;
      if (largura(c, tent) > w && linha.trim()) { out.push(linha.trimEnd()); linha = pal.trimStart(); }
      else linha = tent;
    }
    out.push(linha);
  }
  return out;
}
function desenhaLinhaPlanilha(page, f, r, mapa, cobertas, mesclaDe, cols, alturas, c0, c1, margem, yTopo, H, k, op) {
  const h = alturas[r] * k;
  let x = margem;
  const xs = [];
  for (let c = c0; c < c1; c++) { xs[c] = x; x += cols[c] * k; }
  const xFim = x;
  if (op.grade) {
    for (let c = c0; c < c1; c++) if (cols[c]) {
      const chave = r * 20000 + c;
      if (cobertas.has(chave)) continue;
      const m = mesclaDe.get(chave);
      const w = m ? larguraEntre(cols, c, Math.min(m.c1 + 1, c1)) * k : cols[c] * k;
      const hh = m ? alturaEntre(alturas, r, m.r1 + 1) * k : h;
      page.drawRectangle({ x: xs[c], y: H - yTopo - hh, width: w, height: hh, borderColor: rgb(0.82, 0.82, 0.82), borderWidth: 0.35 });
    }
  }
  for (let c = c0; c < c1; c++) {
    const chave = r * 20000 + c;
    if (cobertas.has(chave) || !cols[c]) continue;
    const cel = mapa.get(chave);
    if (!cel) continue;
    const m = mesclaDe.get(chave);
    const w = m ? larguraEntre(cols, c, Math.min(m.c1 + 1, c1)) * k : cols[c] * k;
    const hh = m ? alturaEntre(alturas, r, m.r1 + 1) * k : h;
    const x0 = xs[c], yB = H - yTopo - hh;
    if (cel.fundo) page.drawRectangle({ x: x0, y: yB, width: w, height: hh, color: cor(cel.fundo) });
    if (cel.borda) {
      const b = cel.borda;
      const lado = (l, a, z) => { if (l) page.drawLine({ start: a, end: z, thickness: l.larg * Math.max(0.6, k), color: cor(l.cor) }); };
      lado(b.t, { x: x0, y: yB + hh }, { x: x0 + w, y: yB + hh });
      lado(b.b, { x: x0, y: yB }, { x: x0 + w, y: yB });
      lado(b.e, { x: x0, y: yB }, { x: x0, y: yB + hh });
      lado(b.d, { x: x0 + w, y: yB }, { x: x0 + w, y: yB + hh });
    }
    if (!cel.texto) continue;
    const r2 = Object.assign({}, cel, { tam: (cel.tam || 11) * k });
    const pad = 2 * k;
    const alin = cel.alin && cel.alin !== 'general' ? cel.alin : (cel.num ? 'right' : 'left');
    let linhas;
    if (cel.quebra) linhas = quebraTexto(r2, Math.max(4, w - 2 * pad));
    else {
      let t = limpa(String(cel.texto).replace(/\r?\n/g, ' '));
      let disp = w - 2 * pad - (cel.recuo || 0) * 9 * k;
      if (alin === 'left' && !m) {
        /* texto pode avançar sobre as células vazias à direita, como no Excel */
        for (let cc = c + 1; cc < c1; cc++) { const vz = mapa.get(r * 20000 + cc); if ((vz && vz.texto) || cobertas.has(r * 20000 + cc)) break; disp += cols[cc] * k; if (largura(r2, t) <= disp) break; }
      }
      if (largura(r2, t) > disp) {
        if (cel.num) t = '#'.repeat(Math.max(1, Math.floor(disp / Math.max(1, largura(r2, '#')))));
        else { while (t.length > 1 && largura(r2, t) > disp) t = t.slice(0, -1); }
      }
      linhas = [t];
    }
    const lh = r2.tam * 1.2;
    const bloco = linhas.length * lh;
    const va = cel.valin || 'bottom';
    let yt = va === 'top' ? 1.5 * k : va === 'center' ? (hh - bloco) / 2 : hh - bloco - 1.5 * k;
    for (const lt of linhas) {
      const lw = largura(r2, lt);
      const xt = alin === 'center' || alin === 'centerContinuous' ? x0 + (w - lw) / 2 : alin === 'right' ? x0 + w - pad - lw : x0 + pad + (cel.recuo || 0) * 9 * k;
      const yb = H - yTopo - yt - r2.tam * 0.95;
      if (lt) page.drawText(lt, { x: xt, y: yb, size: r2.tam, font: fonte(r2), color: cor(cel.cor) });
      if (cel.u && lt) page.drawLine({ start: { x: xt, y: yb - r2.tam * 0.12 }, end: { x: xt + lw, y: yb - r2.tam * 0.12 }, thickness: 0.5, color: cor(cel.cor) });
      yt += lh;
    }
  }
  return xFim;
}
const larguraEntre = (cols, a, b) => { let s = 0; for (let i = a; i < b; i++) s += cols[i] || 0; return s; };
const alturaEntre = (alt, a, b) => { let s = 0; for (let i = a; i < b; i++) s += alt[i] || 0; return s; };

/* ================================================================
   APRESENTAÇÃO (PowerPoint)
   ================================================================ */
function caminhoForma(geo, w, h) {
  const r = (a, b) => `${a.toFixed(2)} ${b.toFixed(2)}`;
  switch (geo) {
    case 'ellipse': case 'flowChartConnector': { const k = 0.5523, rx = w / 2, ry = h / 2; return `M${r(rx, 0)} C${r(rx + k * rx, 0)} ${r(w, ry - k * ry)} ${r(w, ry)} C${r(w, ry + k * ry)} ${r(rx + k * rx, h)} ${r(rx, h)} C${r(rx - k * rx, h)} ${r(0, ry + k * ry)} ${r(0, ry)} C${r(0, ry - k * ry)} ${r(rx - k * rx, 0)} ${r(rx, 0)} Z`; }
    case 'roundRect': case 'flowChartAlternateProcess': { const q = Math.min(w, h) * 0.1667, k = 0.5523 * q; return `M${r(q, 0)} L${r(w - q, 0)} C${r(w - q + k, 0)} ${r(w, q - k)} ${r(w, q)} L${r(w, h - q)} C${r(w, h - q + k)} ${r(w - q + k, h)} ${r(w - q, h)} L${r(q, h)} C${r(q - k, h)} ${r(0, h - q + k)} ${r(0, h - q)} L${r(0, q)} C${r(0, q - k)} ${r(q - k, 0)} ${r(q, 0)} Z`; }
    case 'triangle': case 'flowChartExtract': return `M${r(w / 2, 0)} L${r(w, h)} L${r(0, h)} Z`;
    case 'rtTriangle': return `M${r(0, 0)} L${r(w, h)} L${r(0, h)} Z`;
    case 'diamond': case 'flowChartDecision': return `M${r(w / 2, 0)} L${r(w, h / 2)} L${r(w / 2, h)} L${r(0, h / 2)} Z`;
    case 'parallelogram': case 'flowChartInputOutput': { const a = Math.min(w, h) * 0.25; return `M${r(a, 0)} L${r(w, 0)} L${r(w - a, h)} L${r(0, h)} Z`; }
    case 'trapezoid': { const a = Math.min(w, h) * 0.25; return `M${r(a, 0)} L${r(w - a, 0)} L${r(w, h)} L${r(0, h)} Z`; }
    case 'homePlate': { const a = Math.min(w, h) * 0.5; return `M0 0 L${r(w - a, 0)} L${r(w, h / 2)} L${r(w - a, h)} L${r(0, h)} Z`; }
    case 'chevron': { const a = Math.min(w, h) * 0.5; return `M0 0 L${r(w - a, 0)} L${r(w, h / 2)} L${r(w - a, h)} L${r(0, h)} L${r(a, h / 2)} Z`; }
    case 'pentagon': return `M${r(w / 2, 0)} L${r(w, h * 0.38)} L${r(w * 0.81, h)} L${r(w * 0.19, h)} L${r(0, h * 0.38)} Z`;
    case 'hexagon': { const a = Math.min(w, h) * 0.25; return `M${r(a, 0)} L${r(w - a, 0)} L${r(w, h / 2)} L${r(w - a, h)} L${r(a, h)} L${r(0, h / 2)} Z`; }
    case 'octagon': { const a = Math.min(w, h) * 0.29; return `M${r(a, 0)} L${r(w - a, 0)} L${r(w, a)} L${r(w, h - a)} L${r(w - a, h)} L${r(a, h)} L${r(0, h - a)} L${r(0, a)} Z`; }
    case 'rightArrow': { const a = Math.min(w, h) * 0.5; return `M0 ${r(0, h * 0.25).split(' ')[1]} L${r(w - a, h * 0.25)} L${r(w - a, 0)} L${r(w, h / 2)} L${r(w - a, h)} L${r(w - a, h * 0.75)} L${r(0, h * 0.75)} Z`; }
    case 'leftArrow': { const a = Math.min(w, h) * 0.5; return `M${r(w, h * 0.25)} L${r(a, h * 0.25)} L${r(a, 0)} L${r(0, h / 2)} L${r(a, h)} L${r(a, h * 0.75)} L${r(w, h * 0.75)} Z`; }
    case 'upArrow': { const a = Math.min(w, h) * 0.5; return `M${r(w * 0.25, h)} L${r(w * 0.25, a)} L${r(0, a)} L${r(w / 2, 0)} L${r(w, a)} L${r(w * 0.75, a)} L${r(w * 0.75, h)} Z`; }
    case 'downArrow': { const a = Math.min(w, h) * 0.5; return `M${r(w * 0.25, 0)} L${r(w * 0.75, 0)} L${r(w * 0.75, h - a)} L${r(w, h - a)} L${r(w / 2, h)} L${r(0, h - a)} L${r(w * 0.25, h - a)} Z`; }
    case 'rect': case 'flowChartProcess': case 'snip1Rect': case 'round1Rect': case 'round2SameRect': case 'plaque': case 'frame': case 'bevel': case 'wedgeRectCallout': case 'flowChartPredefinedProcess': case 'flowChartDocument': case 'flowChartTerminator': case 'actionButtonBlank':
      return `M0 0 L${r(w, 0)} L${r(w, h)} L${r(0, h)} Z`;
    default: return null;
  }
}
function comTransformacao(page, f, H, desenha) {
  const precisa = f.rot || f.fh || f.fv;
  if (!precisa) return desenha();
  const cx = f.x + f.w / 2, cy = H - (f.y + f.h / 2);
  const a = -(f.rot || 0) * Math.PI / 180;
  const cos = Math.cos(a), sin = Math.sin(a);
  const sx = f.fh ? -1 : 1, sy = f.fv ? -1 : 1;
  page.pushOperators(pushGraphicsState(), concatTransformationMatrix(1, 0, 0, 1, cx, cy), concatTransformationMatrix(cos * sx, sin * sx, -sin * sy, cos * sy, 0, 0), concatTransformationMatrix(1, 0, 0, 1, -cx, -cy));
  return Promise.resolve(desenha()).then(() => page.pushOperators(popGraphicsState()));
}
async function textoNaCaixa(ctx, tx, x, y, w, h) {
  const ins = tx.ins;
  const bx = x + ins.l, bw = Math.max(4, w - ins.l - ins.r);
  const blocos = [];
  let alt = 0;
  tx.pars.forEach((p, i) => {
    const linhas = quebraParagrafo(p, tx.quebra === false ? 1e5 : bw, { natural: 1.2 });
    /* sem quebra automática: o alinhamento continua relativo à caixa */
    if (tx.quebra === false) for (const L of linhas) L.disp = bw - L.recuo;
    const antes = i ? p.antes : 0;
    alt += antes;
    blocos.push({ p, linhas, antes });
    for (const L of linhas) alt += L.h;
    alt += p.depois || 0;
  });
  const interna = h - ins.t - ins.b;
  let yy = y + ins.t + (tx.ancora === 'ctr' ? (interna - alt) / 2 : tx.ancora === 'b' ? interna - alt : 0);
  for (const b of blocos) {
    yy += b.antes;
    for (const L of b.linhas) { await desenhaLinha(ctx, L, b.p, bx, bw, yy, null); yy += L.h; }
    yy += b.p.depois || 0;
  }
}
async function slidesParaPdf(modelo, id) {
  const doc = await PDFDocument.create();
  await carregaFontes(doc);
  const W = modelo.w, H = modelo.h;
  let k = 0;
  for (const s of modelo.slides) {
    const page = doc.addPage([W, H]);
    const ctx = { page, doc, H };
    if (s.fundo && s.fundo.img) {
      const im = await imagem(doc, s.fundo.img);
      if (im) page.drawImage(im, { x: 0, y: 0, width: W, height: H });
      else page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: rgb(1, 1, 1) });
    } else if (s.fundo && s.fundo.cor) page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: cor(s.fundo.cor) });
    for (const f of s.formas) {
      try {
        if (f.tipo === 'forma') {
          await comTransformacao(page, f, H, async () => {
            if (f.geo === 'linha') {
              const ln = f.ln;
              const x1 = f.x, y1 = f.y, x2 = f.x + f.w, y2 = f.y + f.h;
              if (ln) page.drawLine({ start: { x: x1, y: H - y1 }, end: { x: x2, y: H - y2 }, thickness: Math.max(0.25, ln.larg), color: cor(ln.cor.rgb), opacity: ln.cor.alfa });
            } else {
              const d = f.geo === 'custom' ? f.caminho : caminhoForma(f.geo, f.w, f.h);
              if (d && (f.fill || f.ln)) {
                const opt = { x: f.x, y: H - f.y };
                if (f.fill) { opt.color = cor(f.fill.rgb); opt.opacity = f.fill.alfa; }
                if (f.ln) { opt.borderColor = cor(f.ln.cor.rgb); opt.borderWidth = Math.max(0.25, f.ln.larg); opt.borderOpacity = f.ln.cor.alfa; }
                page.drawSvgPath(d, opt);
              }
            }
            if (f.texto) await textoNaCaixa(ctx, f.texto, f.x, f.y, f.w, f.h);
          });
        } else if (f.tipo === 'imagem') {
          const im = await imagem(doc, f.img);
          if (!im) continue;
          await comTransformacao(page, f, H, () => {
            const c = f.corte;
            if (c && (c.l || c.t || c.r || c.b)) {
              const fw = 1 - c.l - c.r, fh = 1 - c.t - c.b;
              if (fw <= 0 || fh <= 0) return;
              const iw = f.w / fw, ih = f.h / fh;
              page.pushOperators(pushGraphicsState(), rectangle(f.x, H - f.y - f.h, f.w, f.h), clip(), endPath());
              page.drawImage(im, { x: f.x - c.l * iw, y: H - f.y - f.h - c.b * ih, width: iw, height: ih });
              page.pushOperators(popGraphicsState());
            } else page.drawImage(im, { x: f.x, y: H - f.y - f.h, width: f.w, height: f.h });
            if (f.ln) page.drawRectangle({ x: f.x, y: H - f.y - f.h, width: f.w, height: f.h, borderColor: cor(f.ln.cor.rgb), borderWidth: f.ln.larg });
          });
        } else if (f.tipo === 'tabela') await tabelaSlide(ctx, f);
      } catch (e) { /* uma forma com problema não derruba o slide */ }
    }
    k++;
    progresso(id, k, modelo.slides.length);
  }
  return doc;
}
async function tabelaSlide(ctx, t) {
  const { page, H } = ctx;
  const soma = t.cols.reduce((a, c) => a + c, 0) || 1;
  const k = t.w && soma > t.w * 1.02 ? t.w / soma : 1;
  const cols = t.cols.map((c) => c * k);
  const xs = [t.x];
  for (const c of cols) xs.push(xs[xs.length - 1] + c);
  let y = t.y;
  for (const lin of t.linhas) {
    /* altura: a maior entre a definida e a necessária pelo texto */
    let h = lin.alt;
    let c = 0;
    const cels = [];
    for (const cel of lin.celulas) {
      const span = Math.max(1, cel.span || 1);
      const w = (xs[Math.min(xs.length - 1, c + span)] - xs[c]) || cols[c] || 40;
      if (!cel.oculta) {
        let alt = cel.texto.ins.t + cel.texto.ins.b;
        for (const p of cel.texto.pars) { alt += (p.depois || 0); for (const L of quebraParagrafo(p, Math.max(4, w - cel.texto.ins.l - cel.texto.ins.r), { natural: 1.2 })) alt += L.h; }
        if ((cel.rspan || 1) === 1) h = Math.max(h, alt);
      }
      cels.push({ cel, x: xs[c], w });
      c += span;
    }
    for (const { cel, x, w } of cels) {
      if (cel.oculta) continue;
      if (cel.fundo) page.drawRectangle({ x, y: H - y - h, width: w, height: h, color: cor(cel.fundo) });
      page.drawRectangle({ x, y: H - y - h, width: w, height: h, borderColor: rgb(0.75, 0.75, 0.75), borderWidth: 0.6 });
      await textoNaCaixa(ctx, cel.texto, x, y, w, h);
    }
    y += h;
  }
}

/* ================================================================ */
self.onmessage = async (ev) => {
  const { id, op, dados } = ev.data || {};
  trocados = 0; IMGS = {}; cacheImg = new Map();
  try {
    let doc;
    if (dados && dados.modelo && dados.modelo.imagens) IMGS = dados.modelo.imagens;
    if (op === 'documento') doc = await documentoParaPdf(dados.modelo, id);
    else if (op === 'planilhas') doc = await planilhasParaPdf(dados.modelo, dados.opcoes || {}, id);
    else if (op === 'slides') doc = await slidesParaPdf(dados.modelo, id);
    else throw new Falha('INESPERADO');
    doc.setProducer(PRODUTOR);
    doc.setCreator(PRODUTOR);
    if (dados.titulo) doc.setTitle(dados.titulo);
    const bytes = await doc.save({ useObjectStreams: true });
    self.postMessage({ id, tipo: 'ok', resultado: { bytes, paginas: doc.getPageCount(), trocados } }, [bytes.buffer]);
  } catch (e) {
    const memoria = e instanceof RangeError || /allocation|memory|Array buffer/i.test(String(e && e.message));
    self.postMessage({ id, tipo: 'erro', codigo: e instanceof Falha ? e.codigo : memoria ? 'MEMORIA' : 'INESPERADO', detalhe: e instanceof Falha ? e.detalhe : String((e && e.message) || e).slice(0, 200) });
  }
};
