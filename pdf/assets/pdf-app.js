/* GIF Local, ferramentas de PDF.
   Tudo acontece no navegador: os arquivos são lidos do disco do usuário,
   processados aqui e devolvidos como download. Não há nenhuma chamada de
   rede com o conteúdo dos documentos (a CSP destas páginas permite rede
   só para o próprio site). */
/* A versão entra na URL de cada arquivo: o service worker do site guarda
   arquivos estáticos em cache, e URL nova garante código novo após cada
   publicação. Ao alterar qualquer arquivo em /pdf/assets, suba a versão
   aqui e em gerar_paginas.py. */
import { TEXTOS } from './pdf-i18n.js?v=1';
const VERSAO = '1';

const PDFJS_BASE = '/pdf/vendor/pdfjs-4.10.38/';
const IDIOMAS = ['pt', 'en', 'es', 'fr', 'de', 'it'];
const $ = (id) => document.getElementById(id);
const FERRAMENTA = document.body.dataset.ferramenta || 'hub';

/* ---------------- limites ----------------
   Calibrados nos testes (ver relatório): até ~250 MB por arquivo o
   Chromium processou sem travar a interface num computador de 8 GB.
   Em aparelhos com pouca memória informada pelo navegador, o limite cai.
   deviceMemory não existe no Safari/Firefox: nesses casos, celular = 4. */
const MEMORIA_GB = navigator.deviceMemory || (/Android|iPhone|iPad|iPod/i.test(navigator.userAgent) ? 4 : 8);
export const LIMITES = {
  arquivoMB: MEMORIA_GB >= 8 ? 500 : MEMORIA_GB >= 4 ? 250 : 120,
  totalMB: MEMORIA_GB >= 8 ? 800 : MEMORIA_GB >= 4 ? 400 : 200,
  pixelsPorPagina: 16000000, // teto de canvas do Safari no iOS (16,7 Mpx)
};

/* ---------------- idioma ---------------- */
function detectaIdioma() {
  let salvo = null;
  try { salvo = localStorage.getItem('gifLocalLang'); } catch (e) { /* segue sem */ }
  if (salvo && IDIOMAS.includes(salvo)) return salvo;
  const cands = (navigator.languages && navigator.languages.length) ? navigator.languages : [navigator.language || 'pt'];
  for (const c of cands) { const k = String(c).slice(0, 2).toLowerCase(); if (IDIOMAS.includes(k)) return k; }
  return 'en';
}
let idioma = detectaIdioma();
export function t(chave, vars) {
  const d = TEXTOS[idioma] || TEXTOS.pt;
  if (vars && Number(vars.n) === 1 && (d[chave + '_one'] !== undefined || TEXTOS.pt[chave + '_one'] !== undefined)) chave += '_one';
  let s = d[chave] !== undefined ? d[chave] : (TEXTOS.pt[chave] !== undefined ? TEXTOS.pt[chave] : chave);
  if (vars) for (const k of Object.keys(vars)) s = s.split('{' + k + '}').join(String(vars[k]));
  return s;
}
function aplicaTextos() {
  document.documentElement.lang = idioma === 'pt' ? 'pt-BR' : idioma;
  document.querySelectorAll('[data-i18n]').forEach((el) => { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll('[data-i18n-title]').forEach((el) => { el.title = t(el.dataset.i18nTitle); });
  document.querySelectorAll('[data-i18n-aria]').forEach((el) => { el.setAttribute('aria-label', t(el.dataset.i18nAria)); });
  document.querySelectorAll('[data-i18n-limite]').forEach((el) => { el.textContent = t('limit_hint', { mb: LIMITES.arquivoMB }); });
  const chaveTitulo = document.body.dataset.titulo;
  if (chaveTitulo && idioma !== 'pt') document.title = t(chaveTitulo) + ' | GIF Local';
  if (ferramentaAtual && ferramentaAtual.aoMudarIdioma) ferramentaAtual.aoMudarIdioma();
}

/* ---------------- tema ---------------- */
function detectaTema() {
  let s = null;
  try { s = localStorage.getItem('gifLocalTheme'); } catch (e) { /* segue */ }
  if (s === 'light' || s === 'dark') return s;
  return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}
function aplicaTema(tema) {
  document.documentElement.setAttribute('data-theme', tema);
  const sol = $('themeIconSun'), lua = $('themeIconMoon');
  if (sol) sol.hidden = tema === 'dark';
  if (lua) lua.hidden = tema !== 'dark';
  try { localStorage.setItem('gifLocalTheme', tema); } catch (e) { /* segue */ }
}

/* ---------------- formatação ---------------- */
function mb(bytes) { return (bytes / 1048576).toFixed(bytes < 10485760 ? 2 : 1); }
function tamanhoLegivel(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(0) + ' KB';
  return mb(bytes) + ' MB';
}
function paginasTxt(n) { return n === 1 ? t('one_page') : t('n_pages', { n }); }
function nomeBase(nome) { return String(nome || 'documento').replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 80) || 'documento'; }

/* ---------------- worker (pdf-lib) ---------------- */
let worker = null, seq = 0;
const pendentes = new Map();
class ErroFerramenta extends Error {
  constructor(codigo, detalhe) { super(codigo); this.codigo = codigo; this.detalhe = detalhe || null; }
}
function obtemWorker() {
  if (worker) return worker;
  try { worker = new Worker('/pdf/assets/pdf-worker.js?v=' + VERSAO); }
  catch (e) { throw new ErroFerramenta('WORKER'); }
  worker.onmessage = (ev) => {
    const m = ev.data || {};
    const p = pendentes.get(m.id);
    if (!p) return;
    if (m.tipo === 'progresso') { if (p.aoProgredir) p.aoProgredir(m); return; }
    pendentes.delete(m.id);
    if (m.tipo === 'ok') p.resolve(m.resultado);
    else p.reject(new ErroFerramenta(m.codigo, m.detalhe));
  };
  worker.onerror = (e) => {
    e.preventDefault && e.preventDefault();
    /* O worker morreu: por falta de memória (o caso comum com arquivos
       enormes) ou porque o navegador não conseguiu carregá-lo. */
    for (const [, p] of pendentes) p.reject(new ErroFerramenta(p.grande ? 'MEMORIA' : 'WORKER'));
    pendentes.clear();
    try { worker.terminate(); } catch (x) { /* já morto */ }
    worker = null;
  };
  return worker;
}
function chamaWorker(op, dados, transferir, aoProgredir, grande) {
  const w = obtemWorker();
  const id = ++seq;
  return new Promise((resolve, reject) => {
    pendentes.set(id, { resolve, reject, aoProgredir, grande });
    w.postMessage({ id, op, dados }, transferir || []);
  });
}

/* ---------------- PDF.js (renderização) ---------------- */
let pdfjsPromessa = null;
function carregaPdfjs() {
  if (!pdfjsPromessa) {
    pdfjsPromessa = import(PDFJS_BASE + 'pdf.min.mjs').then((m) => {
      m.GlobalWorkerOptions.workerSrc = PDFJS_BASE + 'pdf.worker.min.mjs';
      return m;
    });
  }
  return pdfjsPromessa;
}
async function abrePdfjs(bytes, senha) {
  const pdfjs = await carregaPdfjs();
  /* bytes recém-lidos do arquivo: o PDF.js transfere o buffer para o worker
     dele, sem cópia na thread da página */
  const tarefa = pdfjs.getDocument({
    data: bytes, password: senha, isEvalSupported: false,
    disableAutoFetch: true, disableStream: true, useSystemFonts: true, enableXfa: false,
  });
  try { return await tarefa.promise; }
  catch (e) {
    if (e && e.name === 'PasswordException') throw new ErroFerramenta(e.code === 2 ? 'SENHA_ERRADA' : 'PRECISA_SENHA');
    if (e instanceof RangeError || /memory|allocation/i.test(String(e && e.message))) throw new ErroFerramenta('MEMORIA');
    if (e && /encrypt/i.test(String(e.message))) throw new ErroFerramenta('PROTECAO_INCOMPATIVEL');
    if (e && /Invalid PDF|InvalidPDF|MissingPDF|FormatError|UnknownError/i.test(e.name + ' ' + e.message)) throw new ErroFerramenta('CORROMPIDO');
    throw new ErroFerramenta('CORROMPIDO');
  }
}
async function renderizaPagina(doc, numero, escala, larguraMax) {
  const pg = await doc.getPage(numero);
  let esc = escala;
  if (larguraMax) esc = larguraMax / pg.getViewport({ scale: 1 }).width;
  let vp = pg.getViewport({ scale: esc });
  const area = vp.width * vp.height;
  if (area > LIMITES.pixelsPorPagina) vp = pg.getViewport({ scale: esc * Math.sqrt(LIMITES.pixelsPorPagina / area) });
  const tela = document.createElement('canvas');
  tela.width = Math.max(1, Math.floor(vp.width));
  tela.height = Math.max(1, Math.floor(vp.height));
  const ctx = tela.getContext('2d', { alpha: false });
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, tela.width, tela.height);
  await pg.render({ canvasContext: ctx, viewport: vp }).promise;
  const tamanhoPt = pg.getViewport({ scale: 1 });
  pg.cleanup();
  return { tela, larguraPt: tamanhoPt.width, alturaPt: tamanhoPt.height };
}
function liberaCanvas(c) { if (c) { c.width = 0; c.height = 0; } }
function canvasParaBlob(c, tipo, qualidade) {
  return new Promise((ok, falha) => c.toBlob((b) => (b ? ok(b) : falha(new ErroFerramenta('MEMORIA'))), tipo, qualidade));
}

/* ---------------- validação de arquivos ---------------- */
async function cabecalho(file, n) { return new Uint8Array(await file.slice(0, n).arrayBuffer()); }
async function ehPdf(file) {
  const b = await cabecalho(file, 1024);
  for (let i = 0; i < b.length - 4; i++) {
    if (b[i] === 0x25 && b[i + 1] === 0x50 && b[i + 2] === 0x44 && b[i + 3] === 0x46 && b[i + 4] === 0x2D) return true;
  }
  return false;
}
async function tipoImagem(file) {
  const b = await cabecalho(file, 16);
  if (b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return 'jpg';
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'webp';
  return null;
}
function checaTamanho(file) {
  if (file.size > LIMITES.arquivoMB * 1048576) {
    throw new ErroFerramenta('GRANDE', { nome: file.name, mb: mb(file.size), limite: LIMITES.arquivoMB });
  }
}
async function leBytes(file) {
  try { return new Uint8Array(await file.arrayBuffer()); }
  catch (e) { throw new ErroFerramenta('MEMORIA'); }
}

/* ---------------- ZIP (sem compressão; PNG, JPG e PDF já são comprimidos) ---------------- */
const TABELA_CRC = (() => {
  const tab = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; tab[n] = c >>> 0; }
  return tab;
})();
function crc32(b, c = 0xFFFFFFFF) {
  for (let i = 0; i < b.length; i++) c = TABELA_CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8);
  return c;
}
async function montaZip(itens) {
  const enc = new TextEncoder();
  const partes = [], central = [];
  let desloc = 0;
  const d = new Date();
  const hora = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
  const data = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  const usados = new Set();
  for (const it of itens) {
    let nome = it.nome, k = 2;
    while (usados.has(nome)) nome = it.nome.replace(/(\.[^.]+)?$/, '-' + (k++) + '$1');
    usados.add(nome);
    const nb = enc.encode(nome);
    const bytes = new Uint8Array(await it.blob.arrayBuffer());
    const crc = (crc32(bytes) ^ 0xFFFFFFFF) >>> 0;
    const tam = bytes.length;
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
    lh.setUint16(10, hora, true); lh.setUint16(12, data, true); lh.setUint32(14, crc, true);
    lh.setUint32(18, tam, true); lh.setUint32(22, tam, true); lh.setUint16(26, nb.length, true); lh.setUint16(28, 0, true);
    partes.push(new Uint8Array(lh.buffer), nb, it.blob);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, 0, true); ch.setUint16(12, hora, true); ch.setUint16(14, data, true); ch.setUint32(16, crc, true);
    ch.setUint32(20, tam, true); ch.setUint32(24, tam, true); ch.setUint16(28, nb.length, true);
    ch.setUint32(42, desloc, true);
    central.push(new Uint8Array(ch.buffer), nb);
    desloc += 30 + nb.length + tam;
  }
  let tamCentral = 0;
  for (const c of central) tamCentral += c.length;
  const fim = new DataView(new ArrayBuffer(22));
  fim.setUint32(0, 0x06054b50, true); fim.setUint16(8, itens.length, true); fim.setUint16(10, itens.length, true);
  fim.setUint32(12, tamCentral, true); fim.setUint32(16, desloc, true);
  return new Blob([...partes, ...central, new Uint8Array(fim.buffer)], { type: 'application/zip' });
}

/* ---------------- download ---------------- */
const urlsTemporarias = new Set();
function urlTemporaria(blob) { const u = URL.createObjectURL(blob); urlsTemporarias.add(u); return u; }
function liberaUrls() { for (const u of urlsTemporarias) URL.revokeObjectURL(u); urlsTemporarias.clear(); }
function baixa(blob, nome) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = nome; a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
  estado(t('st_downloaded'), 'ok');
}
function blobPdf(bytes) { return new Blob([bytes], { type: 'application/pdf' }); }

/* ---------------- estado da interface ---------------- */
function estado(texto, tipo) {
  const el = $('status');
  if (!el) return;
  el.textContent = texto || '';
  el.className = 'status' + (tipo ? ' status-' + tipo : '');
  el.setAttribute('role', tipo === 'erro' ? 'alert' : 'status');
}
function progresso(feito, total) {
  const trilha = $('barra'), fill = $('barraFill');
  if (!trilha) return;
  if (feito == null) { trilha.classList.remove('show'); fill.style.width = '0%'; trilha.removeAttribute('aria-valuenow'); return; }
  const pct = total ? Math.round((feito / total) * 100) : 0;
  trilha.classList.add('show');
  trilha.setAttribute('aria-valuenow', String(pct));
  fill.style.width = pct + '%';
}
function mensagemDeErro(e) {
  const c = e && e.codigo;
  const d = (e && e.detalhe) || {};
  switch (c) {
    case 'GRANDE': return t('err_too_big', d);
    case 'NAO_PDF': return t('err_not_pdf', d);
    case 'NAO_IMAGEM': return t('err_not_img', d);
    case 'MEMORIA': return t('err_memory');
    case 'CORROMPIDO': return d.nome ? t('err_corrupt_n', d) : t('err_corrupt');
    case 'CRIPTOGRAFADO': return t('err_encrypted', { nome: d.nome || '' });
    case 'SENHA_ERRADA': return t('err_wrong_pw');
    case 'PRECISA_SENHA': return t('err_need_pw');
    case 'PRECISA_SENHA_PROPRIETARIO': return t('err_need_owner');
    case 'PROTECAO_INCOMPATIVEL': return t('err_unsupported_protection');
    case 'SEM_PROTECAO': return t('unlock_state_open');
    case 'PAGINAS_SINTAXE': return t('err_pages_syntax');
    case 'PAGINAS_FAIXA': return t('err_pages_syntax') + ' ' + t('err_pages_range', d);
    case 'WORKER': return t('err_worker');
    default:
      if (e instanceof RangeError || /memory|allocation/i.test(String(e && e.message))) return t('err_memory');
      return t('err_unexpected');
  }
}
function erro(e) {
  if (!(e instanceof ErroFerramenta)) console.error(e);
  progresso(null);
  estado(mensagemDeErro(e), 'erro');
}

/* ---------------- intervalos de páginas ("1-3, 5, 8") ---------------- */
export function interpretaPaginas(texto, total) {
  const grupos = [];
  const partes = String(texto || '').split(/[,;]+/).map((s) => s.trim()).filter(Boolean);
  if (!partes.length) throw new ErroFerramenta('PAGINAS_SINTAXE');
  for (const p of partes) {
    const m = p.match(/^(\d+)\s*(?:-|–|a|to|bis|à)\s*(\d+)$/i) || p.match(/^(\d+)$/);
    if (!m) throw new ErroFerramenta('PAGINAS_SINTAXE');
    const a = parseInt(m[1], 10), b = m[2] ? parseInt(m[2], 10) : a;
    if (a < 1 || b < 1 || a > total || b > total) throw new ErroFerramenta('PAGINAS_FAIXA', { n: total });
    const g = [];
    if (a <= b) for (let i = a; i <= b; i++) g.push(i - 1);
    else for (let i = a; i >= b; i--) g.push(i - 1);
    grupos.push(g);
  }
  return grupos;
}

/* ---------------- zona de arquivos ---------------- */
function ligaZona(aoReceber) {
  const zona = $('zona'), entrada = $('entrada');
  if (!zona || !entrada) return;
  entrada.addEventListener('change', () => {
    const fs = Array.from(entrada.files || []);
    entrada.value = '';
    if (fs.length) aoReceber(fs);
  });
  ['dragenter', 'dragover'].forEach((ev) => zona.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach((ev) => zona.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.remove('dragover'); }));
  zona.addEventListener('drop', (e) => {
    const fs = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
    if (fs.length) aoReceber(fs);
  });
  /* Soltar fora da zona não pode fazer o navegador abrir o PDF e sair da página. */
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => { if (!zona.contains(e.target)) e.preventDefault(); });
}

function el(tag, attrs, filhos) {
  const n = document.createElement(tag);
  if (attrs) for (const [k, v] of Object.entries(attrs)) {
    if (v == null || v === false) continue;
    if (k === 'class') n.className = v;
    else if (k === 'text') n.textContent = v;
    else if (k.startsWith('on')) n.addEventListener(k.slice(2), v);
    else n.setAttribute(k, v === true ? '' : v);
  }
  if (filhos) for (const f of [].concat(filhos)) if (f) n.appendChild(typeof f === 'string' ? document.createTextNode(f) : f);
  return n;
}

const ICONE_CIMA = '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 19V5M5 12l7-7 7 7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const ICONE_BAIXO = '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 5v14M19 12l-7 7-7-7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const ICONE_X = '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>';
const ICONE_BAIXAR = '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M12 4v11m0 0l-4.5-4.5M12 15l4.5-4.5M5 20h14" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
function botaoIcone(svg, rotulo, aoClicar, desativado) {
  const b = el('button', { type: 'button', class: 'mini-btn', 'aria-label': rotulo, title: rotulo, disabled: !!desativado, onclick: aoClicar });
  b.innerHTML = svg; // SVG fixo definido acima, nunca conteúdo do usuário
  return b;
}
function botaoBaixar(rotulo, aoClicar, classe) {
  const b = el('button', { type: 'button', class: classe || 'btn', onclick: aoClicar });
  b.innerHTML = ICONE_BAIXAR;
  b.appendChild(document.createTextNode(rotulo));
  return b;
}

/* Lista reordenável (setas, teclado e arrastar no desktop). */
function listaReordenavel(container, itens, desenhaItem, aoMudar) {
  container.textContent = '';
  itens.forEach((item, i) => {
    const li = el('li', { class: 'item-arquivo', draggable: 'true' });
    li.dataset.indice = String(i);
    li.appendChild(desenhaItem(item, i));
    const ctl = el('div', { class: 'item-ctl' }, [
      botaoIcone(ICONE_CIMA, t('move_up'), () => { if (i > 0) { [itens[i - 1], itens[i]] = [itens[i], itens[i - 1]]; aoMudar(i - 1); } }, i === 0),
      botaoIcone(ICONE_BAIXO, t('move_down'), () => { if (i < itens.length - 1) { [itens[i + 1], itens[i]] = [itens[i], itens[i + 1]]; aoMudar(i + 1); } }, i === itens.length - 1),
      botaoIcone(ICONE_X, t('remove'), () => { const [r] = itens.splice(i, 1); if (r && r.url) { URL.revokeObjectURL(r.url); urlsTemporarias.delete(r.url); } aoMudar(Math.min(i, itens.length - 1)); }),
    ]);
    li.appendChild(ctl);
    li.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/plain', String(i)); li.classList.add('arrastando'); });
    li.addEventListener('dragend', () => li.classList.remove('arrastando'));
    li.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); });
    li.addEventListener('drop', (e) => {
      e.preventDefault(); e.stopPropagation();
      const de = parseInt(e.dataTransfer.getData('text/plain'), 10);
      if (Number.isInteger(de) && de !== i) { const [m] = itens.splice(de, 1); itens.splice(i, 0, m); aoMudar(i); }
    });
    container.appendChild(li);
  });
}
function focaItem(container, indice) {
  if (indice == null || indice < 0) return;
  const li = container.children[indice];
  const b = li && li.querySelector('.mini-btn:not([disabled])');
  if (b) b.focus();
}

/* ---------------- resultado ---------------- */
function limpaResultado() {
  const r = $('resultado');
  if (r) { r.textContent = ''; r.classList.remove('show'); }
}
function mostraResultado(titulo, estatisticas, acoes, extra) {
  const r = $('resultado');
  r.textContent = '';
  const topo = el('div', { class: 'result-success' });
  const check = el('div', { class: 'success-check', 'aria-hidden': 'true' });
  check.innerHTML = '<svg viewBox="0 0 24 24" fill="none"><path d="M5 12.5l4.5 4.5L19 7.5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  topo.append(check, el('h3', { text: titulo }));
  r.appendChild(topo);
  if (estatisticas && estatisticas.length) {
    const s = el('div', { class: 'result-stats-row' });
    for (const [rot, val] of estatisticas) s.appendChild(el('div', { class: 'stat' }, [el('span', { class: 'stat-label', text: rot }), el('span', { class: 'stat-value', text: val })]));
    r.appendChild(s);
  }
  if (extra) r.appendChild(extra);
  if (acoes && acoes.length) r.appendChild(el('div', { class: 'btn-row result-actions' }, acoes));
  r.classList.add('show');
  const primeiro = r.querySelector('button');
  if (primeiro) primeiro.focus({ preventScroll: false });
}
function listaDeSaidas(saidas) {
  const ul = el('ul', { class: 'saidas' });
  for (const s of saidas) {
    const li = el('li', { class: 'saida' });
    if (s.miniatura) li.appendChild(el('img', { src: s.miniatura, alt: s.nome, loading: 'lazy' }));
    li.appendChild(el('div', { class: 'saida-info' }, [el('b', { text: s.nome }), el('small', { text: [s.detalhe, tamanhoLegivel(s.blob.size)].filter(Boolean).join(' · ') })]));
    li.appendChild(botaoBaixar(t('download_file'), () => baixa(s.blob, s.nome), 'btn secondary small'));
    ul.appendChild(li);
  }
  return ul;
}

/* ---------------- botões principais ---------------- */
function acao(habilitar) { const b = $('acao'); if (b) b.disabled = !habilitar; }
function mostraLimpar(sim) { const b = $('limpar'); if (b) b.hidden = !sim; }
function ocupado(sim) {
  const b = $('acao'); if (b) { b.disabled = sim || b.disabled; b.setAttribute('aria-busy', sim ? 'true' : 'false'); }
  document.querySelectorAll('#opcoes input, #opcoes select, #lista button').forEach((x) => { x.disabled = sim ? true : x.dataset.desativado === '1'; });
}

/* ================================================================
   FERRAMENTAS
   ================================================================ */

/* ---------- Desbloquear ---------- */
function ferramentaDesbloquear() {
  let arq = null, info = null;
  const campo = $('senha'), rotulo = $('senhaRotulo'), aviso = $('protecaoEstado'), mostrar = $('mostrarSenha');
  mostrar.addEventListener('change', () => { campo.type = mostrar.checked ? 'text' : 'password'; });
  campo.addEventListener('input', () => acao(!!info && info.estado !== 'aberto' && campo.value.length > 0));
  campo.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !$('acao').disabled) { e.preventDefault(); processa(); } });

  function desenhaArquivo() {
    const lista = $('lista');
    lista.textContent = '';
    if (!arq) return;
    lista.appendChild(el('li', { class: 'item-arquivo' }, el('div', { class: 'item-info' }, [el('b', { text: arq.name }), el('small', { text: [tamanhoLegivel(arq.size), info && info.paginas ? paginasTxt(info.paginas) : ''].filter(Boolean).join(' · ') })])));
  }
  function textoEstado() {
    if (!info) return;
    aviso.hidden = false;
    aviso.textContent = t(info.estado === 'aberto' ? 'unlock_state_open' : info.estado === 'restrito' ? 'unlock_state_restricted' : 'unlock_state_password');
    rotulo.textContent = t(info.estado === 'restrito' ? 'unlock_pw_owner_label' : 'unlock_pw_label');
  }
  async function recebe(files) {
    limpaResultado();
    const f = files[0];
    try {
      checaTamanho(f);
      estado(t('st_validating'));
      if (!(await ehPdf(f))) throw new ErroFerramenta('NAO_PDF', { nome: f.name });
      arq = f;
      const b = await leBytes(f);
      info = await chamaWorker('info', { bytes: b }, [b.buffer], null, f.size > 50e6);
      desenhaArquivo();
      textoEstado();
      $('opcoes').hidden = info.estado === 'aberto';
      campo.value = '';
      acao(false);
      mostraLimpar(true);
      estado(info.estado === 'aberto' ? t('unlock_state_open') : t('st_selected'), info.estado === 'aberto' ? 'aviso' : null);
      if (info.estado !== 'aberto') campo.focus();
    } catch (e) { arq = null; info = null; desenhaArquivo(); erro(e); }
  }
  async function processa() {
    if (!arq || !info) return;
    const senha = campo.value;
    if (!senha) { estado(t(info.estado === 'restrito' ? 'err_need_owner' : 'err_need_pw'), 'erro'); campo.focus(); return; }
    ocupado(true); limpaResultado();
    estado(t('st_processing'));
    try {
      const b = await leBytes(arq);
      const r = await chamaWorker('desbloquear', { bytes: b, senha }, [b.buffer], (m) => progresso(m.feito, m.total), arq.size > 50e6);
      progresso(null);
      const blob = blobPdf(r.bytes);
      const nome = nomeBase(arq.name) + '-desbloqueado.pdf';
      estado(t('st_done'), 'ok');
      mostraResultado(t('unlock_done'), [[t('lbl_pages'), String(r.paginas)], [t('lbl_size'), tamanhoLegivel(blob.size)]], [botaoBaixar(t('download'), () => baixa(blob, nome))]);
    } catch (e) {
      erro(e);
      if (e.codigo === 'SENHA_ERRADA') { campo.select(); campo.focus(); }
    } finally { ocupado(false); acao(campo.value.length > 0); }
  }
  function limpa() { arq = null; info = null; campo.value = ''; aviso.hidden = true; $('opcoes').hidden = true; desenhaArquivo(); limpaResultado(); acao(false); mostraLimpar(false); progresso(null); estado(t('st_wait')); }
  return { recebe, processa, limpa, aoMudarIdioma: () => { textoEstado(); desenhaArquivo(); } };
}

/* ---------- Imagens para PDF ---------- */
function orientacaoExif(bytes) {
  /* Lê só a tag Orientation (0x0112) do EXIF de um JPEG. */
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (v.byteLength < 4 || v.getUint16(0) !== 0xFFD8) return 1;
  let off = 2;
  while (off + 4 < v.byteLength) {
    const marcador = v.getUint16(off);
    const tam = v.getUint16(off + 2);
    if (marcador === 0xFFE1 && v.getUint32(off + 4) === 0x45786966) {
      const tiff = off + 10;
      const le = v.getUint16(tiff) === 0x4949;
      const ifd = tiff + v.getUint32(tiff + 4, le);
      const n = v.getUint16(ifd, le);
      for (let i = 0; i < n; i++) {
        const e = ifd + 2 + i * 12;
        if (e + 10 > v.byteLength) break;
        if (v.getUint16(e, le) === 0x0112) return v.getUint16(e + 8, le);
      }
      return 1;
    }
    if ((marcador & 0xFF00) !== 0xFF00 || marcador === 0xFFDA) break;
    off += 2 + tam;
  }
  return 1;
}
function ferramentaImagens() {
  const itens = [];
  const lista = $('lista'), previa = $('previa'), previaBox = $('previaBox');
  const opc = () => ({ tamanho: $('tamanho').value, orientacao: $('orientacao').value, margem: $('margem').value, ajuste: $('ajuste').value });
  ['tamanho', 'orientacao', 'margem', 'ajuste'].forEach((id) => $(id).addEventListener('change', desenhaPrevia));

  function desenhaLista(foco) {
    listaReordenavel(lista, itens, (it) => el('div', { class: 'item-info com-mini' }, [
      el('img', { class: 'mini', src: it.url, alt: '', loading: 'lazy' }),
      el('div', {}, [el('b', { text: it.file.name }), el('small', { text: it.largura + ' × ' + it.altura + ' px · ' + tamanhoLegivel(it.file.size) })]),
    ]), (f) => { desenhaLista(f); });
    focaItem(lista, foco);
    desenhaPrevia();
    const n = itens.length;
    acao(n > 0);
    mostraLimpar(n > 0);
    $('opcoes').hidden = n === 0;
    estado(n ? (n === 1 ? t('st_selected') : t('st_selected_n', { n })) : t('st_wait'));
  }
  function desenhaPrevia() {
    previa.textContent = '';
    previaBox.hidden = itens.length === 0;
    const o = opc();
    const dims = { a4: [595.28, 841.89], carta: [612, 792] };
    for (const it of itens.slice(0, 24)) {
      let pw, ph;
      const iw = it.largura, ih = it.altura;
      if (o.tamanho === 'imagem') { pw = iw; ph = ih; }
      else {
        [pw, ph] = dims[o.tamanho];
        if (o.orientacao === 'paisagem' || (o.orientacao === 'auto' && iw > ih)) [pw, ph] = [ph, pw];
      }
      const pagina = el('div', { class: 'pagina-previa' });
      pagina.style.aspectRatio = pw + ' / ' + ph;
      const margem = o.ajuste === 'preencher' && o.tamanho !== 'imagem' ? 0 : o.margem === 'pequena' ? 18 : o.margem === 'grande' ? 42 : 0;
      pagina.style.padding = (margem / pw * 100).toFixed(2) + '%';
      const img = el('img', { src: it.url, alt: '' });
      img.style.objectFit = o.ajuste === 'preencher' ? 'cover' : 'contain';
      pagina.appendChild(img);
      previa.appendChild(pagina);
    }
  }
  async function recebe(files) {
    limpaResultado();
    estado(t('st_validating'));
    for (const f of files) {
      try {
        checaTamanho(f);
        const tipo = await tipoImagem(f);
        if (!tipo) throw new ErroFerramenta('NAO_IMAGEM', { nome: f.name });
        const bmp = await createImageBitmap(f, { imageOrientation: 'from-image' }).catch(() => { throw new ErroFerramenta('NAO_IMAGEM', { nome: f.name }); });
        const it = { file: f, tipo, largura: bmp.width, altura: bmp.height, url: urlTemporaria(f) };
        bmp.close();
        itens.push(it);
      } catch (e) { desenhaLista(); erro(e); return; }
    }
    desenhaLista();
  }
  async function preparaImagem(it) {
    const bytes = await leBytes(it.file);
    if (it.tipo === 'png') return { bytes, tipo: 'png', largura: it.largura, altura: it.altura };
    if (it.tipo === 'jpg' && orientacaoExif(bytes) <= 1) return { bytes, tipo: 'jpg', largura: it.largura, altura: it.altura };
    /* WebP ou JPEG girado pelo EXIF: redesenha já na orientação certa. */
    const bmp = await createImageBitmap(it.file, { imageOrientation: 'from-image' });
    const c = document.createElement('canvas');
    c.width = bmp.width; c.height = bmp.height;
    const ctx = c.getContext('2d');
    ctx.drawImage(bmp, 0, 0);
    bmp.close();
    let transparente = false;
    if (it.tipo === 'webp') {
      const px = ctx.getImageData(0, 0, c.width, c.height).data;
      for (let i = 3; i < px.length; i += 16) if (px[i] < 255) { transparente = true; break; }
    }
    const blob = transparente ? await canvasParaBlob(c, 'image/png') : await canvasParaBlob(c, 'image/jpeg', 0.92);
    liberaCanvas(c);
    return { bytes: new Uint8Array(await blob.arrayBuffer()), tipo: transparente ? 'png' : 'jpg', largura: it.largura, altura: it.altura };
  }
  async function processa() {
    if (!itens.length) return;
    ocupado(true); limpaResultado();
    estado(t('st_processing'));
    try {
      const imagens = [];
      for (let i = 0; i < itens.length; i++) { progresso(i, itens.length * 2); imagens.push(await preparaImagem(itens[i])); }
      const r = await chamaWorker('imagensParaPdf', { imagens, opcoes: opc() }, imagens.map((x) => x.bytes.buffer),
        (m) => progresso(itens.length + m.feito, itens.length * 2), true);
      progresso(null);
      const blob = blobPdf(r.bytes);
      const nome = (itens.length === 1 ? nomeBase(itens[0].file.name) : 'imagens') + '.pdf';
      estado(t('st_done'), 'ok');
      mostraResultado(t('img_done', { n: r.paginas }), [[t('lbl_pages'), String(r.paginas)], [t('lbl_size'), tamanhoLegivel(blob.size)]], [botaoBaixar(t('download'), () => baixa(blob, nome))]);
    } catch (e) { erro(e); } finally { ocupado(false); acao(itens.length > 0); }
  }
  function limpa() { for (const it of itens) URL.revokeObjectURL(it.url); itens.length = 0; liberaUrls(); desenhaLista(); limpaResultado(); progresso(null); }
  return { recebe, processa, limpa, aoMudarIdioma: () => desenhaLista() };
}

/* ---------- PDF para imagens ---------- */
function ferramentaParaImagens() {
  let arq = null, doc = null, senhaUsada;
  const campoPaginas = $('paginas'), escolha = $('paginasModo');
  escolha.addEventListener('change', () => { $('paginasCampo').hidden = escolha.value !== 'escolher'; estima(); });
  ['formato', 'resolucao'].forEach((id) => $(id).addEventListener('change', estima));
  campoPaginas.addEventListener('input', estima);

  function paginasEscolhidas() {
    const n = doc.numPages;
    if (escolha.value !== 'escolher') return Array.from({ length: n }, (_, i) => i);
    const vistos = new Set(), res = [];
    for (const g of interpretaPaginas(campoPaginas.value, n)) for (const i of g) if (!vistos.has(i)) { vistos.add(i); res.push(i); }
    return res;
  }
  function estima() {
    const alvo = $('estimativa');
    if (!doc) { alvo.textContent = ''; return; }
    let n;
    try { n = paginasEscolhidas().length; } catch (e) { alvo.textContent = ''; return; }
    const dpi = parseInt($('resolucao').value, 10);
    /* A4 como referência: bytes por pixel medidos nos testes (PNG ~1,1; JPG ~0,25) */
    const px = (8.27 * dpi) * (11.69 * dpi);
    const fator = $('formato').value === 'png' ? 1.1 : 0.25;
    alvo.textContent = t('toimg_estimate', { mb: Math.max(1, Math.round(n * px * fator / 1048576)) });
  }
  async function abreComSenha(senha) {
    doc = await abrePdfjs(await leBytes(arq), senha);
    senhaUsada = senha;
    $('senhaVer').hidden = true;
    desenhaArquivo();
    $('opcoes').hidden = false;
    acao(true); mostraLimpar(true); estima();
    estado(t('st_selected'));
  }
  function desenhaArquivo() {
    const lista = $('lista');
    lista.textContent = '';
    if (!arq) return;
    lista.appendChild(el('li', { class: 'item-arquivo' }, el('div', { class: 'item-info' }, [el('b', { text: arq.name }), el('small', { text: [tamanhoLegivel(arq.size), doc ? paginasTxt(doc.numPages) : ''].filter(Boolean).join(' · ') })])));
  }
  async function recebe(files) {
    limpaResultado(); await fecha();
    const f = files[0];
    try {
      checaTamanho(f);
      estado(t('st_validating'));
      if (!(await ehPdf(f))) throw new ErroFerramenta('NAO_PDF', { nome: f.name });
      arq = f;
      desenhaArquivo();
      try { await abreComSenha(undefined); }
      catch (e) {
        if (e.codigo === 'PRECISA_SENHA' || e.codigo === 'SENHA_ERRADA') { $('senhaVer').hidden = false; $('senhaVerCampo').focus(); estado(t('pw_view_label')); mostraLimpar(true); return; }
        throw e;
      }
    } catch (e) { arq = null; desenhaArquivo(); erro(e); }
  }
  $('senhaVerBtn').addEventListener('click', async () => {
    try { await abreComSenha($('senhaVerCampo').value); } catch (e) { erro(e); }
  });
  $('senhaVerCampo').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); $('senhaVerBtn').click(); } });

  async function processa() {
    if (!doc) return;
    let paginas;
    try { paginas = paginasEscolhidas(); } catch (e) { erro(e); return; }
    ocupado(true); limpaResultado(); liberaUrls();
    estado(t('st_processing'));
    const png = $('formato').value === 'png';
    const dpi = parseInt($('resolucao').value, 10);
    const saidas = [];
    try {
      for (let k = 0; k < paginas.length; k++) {
        progresso(k, paginas.length);
        const { tela } = await renderizaPagina(doc, paginas[k] + 1, dpi / 72);
        const blob = await canvasParaBlob(tela, png ? 'image/png' : 'image/jpeg', png ? undefined : 0.9);
        const detalhe = tela.width + ' × ' + tela.height + ' px';
        liberaCanvas(tela);
        const nome = nomeBase(arq.name) + '-pagina-' + String(paginas[k] + 1).padStart(String(doc.numPages).length, '0') + (png ? '.png' : '.jpg');
        saidas.push({ nome, blob, detalhe, miniatura: urlTemporaria(blob) });
      }
      progresso(null);
      estado(t('st_done'), 'ok');
      const acoes = [];
      if (saidas.length > 1) acoes.push(botaoBaixar(t('download_all_zip'), async () => { estado(t('st_processing')); baixa(await montaZip(saidas), nomeBase(arq.name) + '-imagens.zip'); }));
      else acoes.push(botaoBaixar(t('download_file'), () => baixa(saidas[0].blob, saidas[0].nome)));
      const total = saidas.reduce((a, s) => a + s.blob.size, 0);
      mostraResultado(t('toimg_done', { n: saidas.length }), [[t('lbl_files'), String(saidas.length)], [t('lbl_size'), tamanhoLegivel(total)]], acoes, listaDeSaidas(saidas));
    } catch (e) { erro(e); } finally { ocupado(false); acao(!!doc); }
  }
  async function fecha() { if (doc) { try { await doc.destroy(); } catch (e) { /* já fechado */ } } doc = null; }
  async function limpa() { await fecha(); arq = null; senhaUsada = undefined; $('senhaVer').hidden = true; $('opcoes').hidden = true; desenhaArquivo(); limpaResultado(); liberaUrls(); acao(false); mostraLimpar(false); progresso(null); estado(t('st_wait')); }
  return { recebe, processa, limpa, aoMudarIdioma: () => { desenhaArquivo(); estima(); } };
}

/* ---------- Unir ---------- */
function ferramentaUnir() {
  const itens = [];
  const lista = $('lista');
  async function miniatura(it) {
    try {
      const bytes = await leBytes(it.file);
      const doc = await abrePdfjs(bytes);
      it.paginas = doc.numPages;
      const { tela } = await renderizaPagina(doc, 1, null, 96);
      it.url = urlTemporaria(await canvasParaBlob(tela, 'image/jpeg', 0.8));
      liberaCanvas(tela);
      await doc.destroy();
    } catch (e) {
      if (e.codigo === 'PRECISA_SENHA' || e.codigo === 'SENHA_ERRADA') it.problema = 'CRIPTOGRAFADO';
      else it.problema = 'CORROMPIDO';
    }
  }
  function desenha(foco) {
    listaReordenavel(lista, itens, (it) => el('div', { class: 'item-info com-mini' }, [
      it.url ? el('img', { class: 'mini mini-pagina', src: it.url, alt: '' }) : el('span', { class: 'mini mini-vazia', 'aria-hidden': 'true', text: 'PDF' }),
      el('div', {}, [
        el('b', { text: it.file.name }),
        el('small', { class: it.problema ? 'texto-erro' : '', text: it.problema ? mensagemDeErro(new ErroFerramenta(it.problema, { nome: it.file.name })) : [tamanhoLegivel(it.file.size), it.paginas ? paginasTxt(it.paginas) : '…'].join(' · ') }),
      ]),
    ]), (f) => desenha(f));
    focaItem(lista, foco);
    const validos = itens.filter((x) => !x.problema).length;
    const temProblema = itens.some((x) => x.problema);
    acao(validos >= 2 && !temProblema);
    mostraLimpar(itens.length > 0);
    $('adicionarMais').hidden = itens.length === 0;
    if (!itens.length) estado(t('st_wait'));
    else if (temProblema) estado(mensagemDeErro(new ErroFerramenta(itens.find((x) => x.problema).problema, { nome: itens.find((x) => x.problema).file.name })), 'erro');
    else if (validos < 2) estado(t('merge_need2'), 'aviso');
    else estado(t('st_selected_n', { n: validos }));
  }
  async function recebe(files) {
    limpaResultado();
    estado(t('st_validating'));
    let total = itens.reduce((a, x) => a + x.file.size, 0);
    for (const f of files) {
      try {
        checaTamanho(f);
        total += f.size;
        if (total > LIMITES.totalMB * 1048576) throw new ErroFerramenta('GRANDE', { nome: f.name, mb: mb(total), limite: LIMITES.totalMB });
        if (!(await ehPdf(f))) throw new ErroFerramenta('NAO_PDF', { nome: f.name });
        const it = { file: f };
        itens.push(it);
        desenha();
        await miniatura(it);
        desenha();
      } catch (e) { erro(e); return; }
    }
  }
  async function processa() {
    if (itens.length < 2) return;
    ocupado(true); limpaResultado();
    estado(t('st_processing'));
    try {
      const arquivos = [];
      for (const it of itens) arquivos.push(await leBytes(it.file));
      const tamanho = arquivos.reduce((a, b) => a + b.length, 0);
      const r = await chamaWorker('unir', { arquivos }, arquivos.map((b) => b.buffer), (m) => progresso(m.feito, m.total), tamanho > 80e6).catch((e) => {
        if (e.codigo === 'CRIPTOGRAFADO' || e.codigo === 'CORROMPIDO') { const it = itens[(e.detalhe && e.detalhe.indice) || 0]; e.detalhe = { nome: it ? it.file.name : '' }; }
        throw e;
      });
      progresso(null);
      const blob = blobPdf(r.bytes);
      estado(t('st_done'), 'ok');
      mostraResultado(t('merge_done', { n: r.paginas }), [[t('lbl_files'), String(itens.length)], [t('lbl_pages'), String(r.paginas)], [t('lbl_size'), tamanhoLegivel(blob.size)]],
        [botaoBaixar(t('download'), () => baixa(blob, 'unido.pdf'))]);
    } catch (e) { erro(e); } finally { ocupado(false); desenhaEstadoBotao(); }
  }
  function desenhaEstadoBotao() { acao(itens.filter((x) => !x.problema).length >= 2 && !itens.some((x) => x.problema)); }
  function limpa() { for (const it of itens) if (it.url) URL.revokeObjectURL(it.url); itens.length = 0; liberaUrls(); desenha(); limpaResultado(); progresso(null); }
  $('adicionarMais').addEventListener('click', () => $('entrada').click());
  return { recebe, processa, limpa, aoMudarIdioma: () => desenha() };
}

/* ---------- Dividir ---------- */
function ferramentaDividir() {
  let arq = null, doc = null;
  const selecionadas = new Set();
  const modo = $('modo'), grade = $('grade'), campo = $('intervalos');
  let observador = null;
  modo.addEventListener('change', atualizaModo);
  campo.addEventListener('input', atualizaBotao);

  function atualizaModo() {
    $('campoIntervalos').hidden = modo.value !== 'intervalos';
    $('campoSelecao').hidden = modo.value !== 'selecionar';
    if (modo.value === 'selecionar') montaGrade();
    atualizaBotao();
  }
  function atualizaBotao() {
    if (!doc) { acao(false); return; }
    if (modo.value === 'selecionar') { acao(selecionadas.size > 0); $('contagemSelecao').textContent = t('split_selected', { n: selecionadas.size }); }
    else if (modo.value === 'intervalos') acao(campo.value.trim().length > 0);
    else acao(true);
  }
  function montaGrade() {
    if (!doc || grade.dataset.montada === '1') return;
    grade.dataset.montada = '1';
    grade.textContent = '';
    if (observador) observador.disconnect();
    observador = new IntersectionObserver(async (entradas) => {
      for (const en of entradas) {
        if (!en.isIntersecting) continue;
        const b = en.target;
        observador.unobserve(b);
        try {
          const { tela } = await renderizaPagina(doc, parseInt(b.dataset.pagina, 10), null, 110);
          const u = urlTemporaria(await canvasParaBlob(tela, 'image/jpeg', 0.75));
          liberaCanvas(tela);
          const img = b.querySelector('img');
          img.src = u;
        } catch (e) { /* miniatura é só apoio visual */ }
      }
    }, { rootMargin: '300px' });
    for (let i = 1; i <= doc.numPages; i++) {
      const b = el('button', { type: 'button', class: 'pagina-sel', 'aria-pressed': selecionadas.has(i - 1) ? 'true' : 'false', 'aria-label': t('page_n', { n: i }) },
        [el('img', { alt: '' }), el('span', { class: 'num', text: String(i) })]);
      b.dataset.pagina = String(i);
      b.addEventListener('click', () => {
        const idx = i - 1;
        if (selecionadas.has(idx)) selecionadas.delete(idx); else selecionadas.add(idx);
        b.setAttribute('aria-pressed', selecionadas.has(idx) ? 'true' : 'false');
        atualizaBotao();
      });
      grade.appendChild(b);
      observador.observe(b);
    }
  }
  function desenhaArquivo() {
    const lista = $('lista');
    lista.textContent = '';
    if (!arq) return;
    lista.appendChild(el('li', { class: 'item-arquivo' }, el('div', { class: 'item-info' }, [el('b', { text: arq.name }), el('small', { text: [tamanhoLegivel(arq.size), doc ? paginasTxt(doc.numPages) : ''].filter(Boolean).join(' · ') })])));
  }
  async function recebe(files) {
    await limpa(true);
    const f = files[0];
    try {
      checaTamanho(f);
      estado(t('st_validating'));
      if (!(await ehPdf(f))) throw new ErroFerramenta('NAO_PDF', { nome: f.name });
      arq = f;
      try { doc = await abrePdfjs(await leBytes(f)); }
      catch (e) { if (e.codigo === 'PRECISA_SENHA' || e.codigo === 'SENHA_ERRADA') throw new ErroFerramenta('CRIPTOGRAFADO', { nome: f.name }); throw e; }
      desenhaArquivo();
      $('opcoes').hidden = false;
      mostraLimpar(true);
      campo.value = doc.numPages > 1 ? '1-' + Math.ceil(doc.numPages / 2) + ', ' + (Math.ceil(doc.numPages / 2) + 1) + '-' + doc.numPages : '1';
      atualizaModo();
      estado(t('st_selected'));
    } catch (e) { arq = null; desenhaArquivo(); erro(e); }
  }
  async function processa() {
    if (!doc) return;
    let grupos, rotulos;
    const n = doc.numPages;
    try {
      if (modo.value === 'cada') { grupos = Array.from({ length: n }, (_, i) => [i]); rotulos = grupos.map((g) => 'pagina-' + (g[0] + 1)); }
      else if (modo.value === 'selecionar') {
        if (!selecionadas.size) { estado(t('split_none'), 'erro'); return; }
        grupos = [Array.from(selecionadas).sort((a, b) => a - b)]; rotulos = ['paginas-selecionadas'];
      } else {
        grupos = interpretaPaginas(campo.value, n);
        rotulos = grupos.map((g) => g.length === 1 ? 'pagina-' + (g[0] + 1) : 'paginas-' + (g[0] + 1) + '-' + (g[g.length - 1] + 1));
      }
    } catch (e) { erro(e); return; }
    ocupado(true); limpaResultado();
    estado(t('st_processing'));
    try {
      const b = await leBytes(arq);
      const r = await chamaWorker('dividir', { bytes: b, grupos, rotulos }, [b.buffer], (m) => progresso(m.feito, m.total), arq.size > 50e6);
      progresso(null);
      const base = nomeBase(arq.name);
      const saidas = r.saidas.map((s) => ({ nome: base + '-' + s.rotulo + '.pdf', blob: blobPdf(s.bytes), detalhe: paginasTxt(s.paginas) }));
      estado(t('st_done'), 'ok');
      const acoes = saidas.length > 1
        ? [botaoBaixar(t('download_all_zip'), async () => { estado(t('st_processing')); baixa(await montaZip(saidas), base + '-dividido.zip'); })]
        : [botaoBaixar(t('download'), () => baixa(saidas[0].blob, saidas[0].nome))];
      mostraResultado(t('split_done', { n: saidas.length }), [[t('lbl_files'), String(saidas.length)]], acoes, saidas.length > 1 ? listaDeSaidas(saidas) : null);
    } catch (e) { erro(e); } finally { ocupado(false); atualizaBotao(); }
  }
  async function limpa(silencioso) {
    if (observador) observador.disconnect();
    if (doc) { try { await doc.destroy(); } catch (e) { /* fechado */ } }
    doc = null; arq = null; selecionadas.clear();
    grade.textContent = ''; grade.dataset.montada = '';
    liberaUrls(); desenhaArquivo(); limpaResultado(); progresso(null);
    $('opcoes').hidden = true; acao(false); mostraLimpar(false);
    if (!silencioso) estado(t('st_wait'));
  }
  return { recebe, processa, limpa: () => limpa(false), aoMudarIdioma: () => { desenhaArquivo(); atualizaBotao(); } };
}

/* ---------- Comprimir ---------- */
function ferramentaComprimir() {
  let arq = null, paginas = 0;
  function desenhaArquivo() {
    const lista = $('lista');
    lista.textContent = '';
    if (!arq) return;
    lista.appendChild(el('li', { class: 'item-arquivo' }, el('div', { class: 'item-info' }, [el('b', { text: arq.name }), el('small', { text: [tamanhoLegivel(arq.size), paginas ? paginasTxt(paginas) : ''].filter(Boolean).join(' · ') })])));
  }
  async function recebe(files) {
    limpaResultado();
    const f = files[0];
    try {
      checaTamanho(f);
      estado(t('st_validating'));
      if (!(await ehPdf(f))) throw new ErroFerramenta('NAO_PDF', { nome: f.name });
      arq = f;
      const b = await leBytes(f);
      const info = await chamaWorker('info', { bytes: b }, [b.buffer], null, f.size > 50e6);
      if (info.estado !== 'aberto') throw new ErroFerramenta('CRIPTOGRAFADO', { nome: f.name });
      paginas = info.paginas;
      desenhaArquivo();
      $('opcoes').hidden = false;
      acao(true); mostraLimpar(true);
      estado(t('st_selected'));
    } catch (e) { arq = null; paginas = 0; desenhaArquivo(); erro(e); }
  }
  function nivel() { const r = document.querySelector('input[name="nivel"]:checked'); return r ? r.value : 'media'; }
  async function maxima() {
    /* Cada página vira uma imagem JPEG de 110 ppp com o mesmo tamanho físico. */
    const doc = await abrePdfjs(await leBytes(arq));
    const imagens = [];
    try {
      for (let i = 1; i <= doc.numPages; i++) {
        progresso(i - 1, doc.numPages * 1.1);
        const { tela, larguraPt, alturaPt } = await renderizaPagina(doc, i, 110 / 72);
        const blob = await canvasParaBlob(tela, 'image/jpeg', 0.6);
        imagens.push({ bytes: new Uint8Array(await blob.arrayBuffer()), tipo: 'jpg', pagina: [larguraPt, alturaPt] });
        liberaCanvas(tela);
      }
    } finally { await doc.destroy(); }
    return chamaWorker('imagensParaPdf', { imagens }, imagens.map((x) => x.bytes.buffer), null, true);
  }
  async function processa() {
    if (!arq) return;
    ocupado(true); limpaResultado();
    estado(t('st_processing'));
    try {
      const n = nivel();
      const r = n === 'maxima' ? await maxima()
        : await (async () => { const b = await leBytes(arq); return chamaWorker('comprimir', { bytes: b, nivel: n }, [b.buffer], (m) => progresso(m.feito, m.total), arq.size > 50e6); })();
      progresso(null);
      const antes = arq.size, depois = r.bytes.length;
      const reducao = (1 - depois / antes) * 100;
      const stats = [[t('comp_original'), tamanhoLegivel(antes)], [t('comp_final'), tamanhoLegivel(depois)], [t('comp_reduction'), (reducao > 0 ? '−' : '+') + Math.abs(reducao).toFixed(reducao > -1 && reducao < 1 ? 1 : 0) + '%']];
      const blob = blobPdf(r.bytes);
      const nome = nomeBase(arq.name) + '-comprimido.pdf';
      let aviso = null;
      if (r.stats && r.stats.semSuporteNavegador) aviso = el('p', { class: 'note', text: t('comp_no_offscreen') });
      if (depois >= antes) {
        estado(t('comp_bigger'), 'aviso');
        mostraResultado(t('comp_bigger'), stats, [], aviso);
      } else {
        const pouco = reducao < 5;
        estado(pouco ? t('comp_no_gain') : t('st_done'), pouco ? 'aviso' : 'ok');
        mostraResultado(pouco ? t('comp_no_gain') : t('st_done'), stats, [botaoBaixar(t('download'), () => baixa(blob, nome))], aviso);
      }
    } catch (e) { erro(e); } finally { ocupado(false); acao(!!arq); }
  }
  function limpa() { arq = null; paginas = 0; desenhaArquivo(); limpaResultado(); progresso(null); $('opcoes').hidden = true; acao(false); mostraLimpar(false); estado(t('st_wait')); }
  return { recebe, processa, limpa, aoMudarIdioma: desenhaArquivo };
}

/* ================================================================ */
const FABRICAS = {
  desbloquear: ferramentaDesbloquear,
  'imagem-para-pdf': ferramentaImagens,
  'pdf-para-imagem': ferramentaParaImagens,
  unir: ferramentaUnir,
  dividir: ferramentaDividir,
  comprimir: ferramentaComprimir,
};
let ferramentaAtual = null;

function inicia() {
  aplicaTema(detectaTema());
  const tt = $('themeToggle');
  if (tt) tt.addEventListener('click', () => aplicaTema(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'));
  const sel = $('langSwitch');
  if (sel) {
    sel.value = idioma;
    sel.addEventListener('change', () => {
      idioma = IDIOMAS.includes(sel.value) ? sel.value : 'pt';
      try { localStorage.setItem('gifLocalLang', idioma); } catch (e) { /* segue */ }
      aplicaTextos();
    });
  }
  const fabrica = FABRICAS[FERRAMENTA];
  if (fabrica) {
    ferramentaAtual = fabrica();
    ligaZona((fs) => ferramentaAtual.recebe(fs));
    $('acao').addEventListener('click', () => ferramentaAtual.processa());
    $('limpar').addEventListener('click', () => { ferramentaAtual.limpa(); const e = $('entrada'); if (e) e.focus(); });
    estado(t('st_wait'));
  }
  aplicaTextos();
  if (fabrica && !$('status').textContent) estado(t('st_wait'));
  document.documentElement.classList.add('js-pronto');
}

/* Exposto só para os testes automatizados (não altera nada). */
window.__pdfTools = { interpretaPaginas, LIMITES, montaZip, crc32: (b) => (crc32(b) ^ 0xFFFFFFFF) >>> 0 };

inicia();
