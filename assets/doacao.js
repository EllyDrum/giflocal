/* GIF Local, apoio voluntário internacional.
   Idioma, país, moeda, cotação, valores e formas de pagamento num módulo
   só. Regras:
   - valores guardados SEMPRE em unidade mínima inteira (centavos; iene e
     peso chileno sem decimais), nunca em número decimal;
   - a cotação é informativa (fonte e data mostradas); quem define a
     cobrança é o gateway, e a moeda pedida nunca muda em silêncio;
   - só aparecem formas de pagamento que existem de verdade: PIX (BRL,
     gerado aqui) e cartão via Stripe Checkout (qualquer moeda aceita). */
(function () {
  'use strict';

  const API = 'https://api.giflocal.com';
  const MOEDAS = {
    BRL: { dec: 2, locale: 'pt-BR', sugeridos: [500, 1000, 2000, 5000] },
    USD: { dec: 2, locale: 'en-US', sugeridos: [200, 500, 1000, 2000] },
    EUR: { dec: 2, locale: 'de-DE', sugeridos: [200, 500, 1000, 2000] },
    GBP: { dec: 2, locale: 'en-GB', sugeridos: [200, 500, 1000, 2000] },
    CAD: { dec: 2, locale: 'en-CA', sugeridos: [300, 500, 1000, 2500] },
    AUD: { dec: 2, locale: 'en-AU', sugeridos: [300, 500, 1000, 2500] },
    CHF: { dec: 2, locale: 'de-CH', sugeridos: [200, 500, 1000, 2000] },
    JPY: { dec: 0, locale: 'ja-JP', sugeridos: [500, 1000, 3000, 5000, 10000] },
    MXN: { dec: 2, locale: 'es-MX', sugeridos: [5000, 10000, 20000, 50000] },
    ARS: { dec: 2, locale: 'es-AR', sugeridos: [200000, 500000, 1000000, 2000000] },
    CLP: { dec: 0, locale: 'es-CL', sugeridos: [2000, 5000, 10000, 20000] },
  };
  /* limites locais (espelham o Worker; o Worker é quem manda) */
  const LIMITES = { BRL: [100, 2000000], USD: [50, 500000], EUR: [50, 500000], GBP: [30, 400000], CAD: [50, 700000], AUD: [50, 750000], CHF: [50, 450000], JPY: [50, 750000], MXN: [1000, 9000000], ARS: [50, 99999999], CLP: [500, 5000000] };
  const NOMES = {
    pt: { BRL: 'Real brasileiro', USD: 'Dólar americano', EUR: 'Euro', GBP: 'Libra esterlina', CAD: 'Dólar canadense', AUD: 'Dólar australiano', CHF: 'Franco suíço', JPY: 'Iene japonês', MXN: 'Peso mexicano', ARS: 'Peso argentino', CLP: 'Peso chileno' },
    en: { BRL: 'Brazilian real', USD: 'US dollar', EUR: 'Euro', GBP: 'British pound', CAD: 'Canadian dollar', AUD: 'Australian dollar', CHF: 'Swiss franc', JPY: 'Japanese yen', MXN: 'Mexican peso', ARS: 'Argentine peso', CLP: 'Chilean peso' },
    es: { BRL: 'Real brasileño', USD: 'Dólar estadounidense', EUR: 'Euro', GBP: 'Libra esterlina', CAD: 'Dólar canadiense', AUD: 'Dólar australiano', CHF: 'Franco suizo', JPY: 'Yen japonés', MXN: 'Peso mexicano', ARS: 'Peso argentino', CLP: 'Peso chileno' },
    fr: { BRL: 'Réal brésilien', USD: 'Dollar américain', EUR: 'Euro', GBP: 'Livre sterling', CAD: 'Dollar canadien', AUD: 'Dollar australien', CHF: 'Franc suisse', JPY: 'Yen japonais', MXN: 'Peso mexicain', ARS: 'Peso argentin', CLP: 'Peso chilien' },
    de: { BRL: 'Brasilianischer Real', USD: 'US-Dollar', EUR: 'Euro', GBP: 'Britisches Pfund', CAD: 'Kanadischer Dollar', AUD: 'Australischer Dollar', CHF: 'Schweizer Franken', JPY: 'Japanischer Yen', MXN: 'Mexikanischer Peso', ARS: 'Argentinischer Peso', CLP: 'Chilenischer Peso' },
    it: { BRL: 'Real brasiliano', USD: 'Dollaro statunitense', EUR: 'Euro', GBP: 'Sterlina britannica', CAD: 'Dollaro canadese', AUD: 'Dollaro australiano', CHF: 'Franco svizzero', JPY: 'Yen giapponese', MXN: 'Peso messicano', ARS: 'Peso argentino', CLP: 'Peso cileno' },
    ja: { BRL: 'ブラジルレアル', USD: '米ドル', EUR: 'ユーロ', GBP: '英ポンド', CAD: 'カナダドル', AUD: 'オーストラリアドル', CHF: 'スイスフラン', JPY: '日本円', MXN: 'メキシコペソ', ARS: 'アルゼンチンペソ', CLP: 'チリペソ' },
  };
  /* moeda inicial: região do navegador primeiro, depois o idioma */
  const MOEDA_POR_REGIAO = { BR: 'BRL', US: 'USD', GB: 'GBP', CA: 'CAD', AU: 'AUD', CH: 'CHF', JP: 'JPY', MX: 'MXN', AR: 'ARS', CL: 'CLP', PT: 'EUR', ES: 'EUR', FR: 'EUR', DE: 'EUR', IT: 'EUR', AT: 'EUR', BE: 'EUR', NL: 'EUR', IE: 'EUR', FI: 'EUR', GR: 'EUR' };
  const MOEDA_POR_IDIOMA = { pt: 'BRL', en: 'USD', es: 'EUR', fr: 'EUR', de: 'EUR', it: 'EUR', ja: 'JPY' };
  const LOCALE_IDIOMA = { pt: 'pt-BR', en: 'en-US', es: 'es-ES', fr: 'fr-FR', de: 'de-DE', it: 'it-IT', ja: 'ja-JP' };
  const CHAVE_MOEDA = 'gifLocalMoeda';

  /* ---------- formatação ---------- */
  function locale(idioma, moeda) {
    /* o idioma manda na pontuação; a moeda só define símbolo e casas */
    return LOCALE_IDIOMA[idioma] || (MOEDAS[moeda] && MOEDAS[moeda].locale) || 'en-US';
  }
  function formata(minor, moeda, idioma) {
    const m = MOEDAS[moeda]; if (!m) return String(minor);
    const v = minor / Math.pow(10, m.dec);
    try { return new Intl.NumberFormat(locale(idioma, moeda), { style: 'currency', currency: moeda, currencyDisplay: 'narrowSymbol', minimumFractionDigits: m.dec, maximumFractionDigits: m.dec }).format(v).replace(/\uFFE5/g, '\u00A5'); }
    catch (e1) { try { return new Intl.NumberFormat(locale(idioma, moeda), { style: 'currency', currency: moeda, minimumFractionDigits: m.dec, maximumFractionDigits: m.dec }).format(v); } catch (e2) { /* abaixo */ } }
    return moeda + ' ' + v.toFixed(m.dec);
  }
  function formataComCodigo(minor, moeda, idioma) { return formata(minor, moeda, idioma) + ' ' + moeda; }
  function separadores(idioma, moeda) {
    try {
      const partes = new Intl.NumberFormat(locale(idioma, moeda)).formatToParts(1234567.891);
      return { dec: (partes.find((p) => p.type === 'decimal') || {}).value || '.', mil: (partes.find((p) => p.type === 'group') || {}).value || ',' };
    } catch (e) { return { dec: '.', mil: ',' }; }
  }
  /* Texto digitado → unidade mínima inteira (ou null). Aceita "35,50",
     "35.50", "1.000,00", "1,000.00", "¥3,500", "R$ 35,50". Sem ponto
     flutuante: só manipulação de dígitos. */
  function interpreta(texto, moeda, idioma) {
    const m = MOEDAS[moeda]; if (!m) return null;
    let s = String(texto || '').replace(/[^\d.,٫٬]/g, '').replace(/٫/g, ',').replace(/٬/g, '.');
    if (!s) return null;
    const sep = separadores(idioma, moeda);
    const pontos = (s.match(/\./g) || []).length, virgulas = (s.match(/,/g) || []).length;
    let inteiro = s, fracao = '';
    if (m.dec === 0) {
      inteiro = s.replace(/[.,]/g, '');
      /* "3.5" ou "3,5" numa moeda sem decimais: provavelmente erro de digitação */
      const ult = s.search(/[.,][^.,]*$/);
      if (ult >= 0 && s.length - ult - 1 !== 3 && pontos + virgulas === 1) return null;
    } else if (pontos + virgulas === 0) {
      inteiro = s;
    } else if (pontos && virgulas) {
      const ultimo = Math.max(s.lastIndexOf('.'), s.lastIndexOf(','));
      inteiro = s.slice(0, ultimo).replace(/[.,]/g, ''); fracao = s.slice(ultimo + 1);
      if (!/^\d{1,2}$/.test(fracao)) return null;
    } else {
      const ch = pontos ? '.' : ',';
      const partes = s.split(ch);
      const resto = partes[partes.length - 1];
      if (partes.length === 2 && resto.length <= 2) { inteiro = partes[0]; fracao = resto; }
      else if (partes.length === 2 && resto.length === 3 && ch === sep.mil) { inteiro = partes.join(''); }
      else if (partes.length === 2 && resto.length === 3 && ch === sep.dec) { return null; }
      else if (partes.every((p, i) => (i === 0 ? p.length <= 3 : p.length === 3))) { inteiro = partes.join(''); }
      else return null;
    }
    if (!/^\d+$/.test(inteiro) || inteiro.length > 12) return null;
    const minor = parseInt(inteiro + (m.dec ? (fracao + '00').slice(0, m.dec) : ''), 10);
    return Number.isSafeInteger(minor) && minor > 0 ? minor : null;
  }
  /* conversão informativa a partir das taxas base BRL (sem flutuante nos extremos: arredonda só no fim) */
  function converte(minor, de, para, taxas) {
    if (!taxas || !taxas[de] || !taxas[para] || !MOEDAS[de] || !MOEDAS[para]) return null;
    const valorDe = minor / Math.pow(10, MOEDAS[de].dec);
    const valorPara = valorDe / taxas[de] * taxas[para];
    return Math.round(valorPara * Math.pow(10, MOEDAS[para].dec));
  }
  function moedaInicial(idioma) {
    try { const s = localStorage.getItem(CHAVE_MOEDA); if (s && MOEDAS[s]) return s; } catch (e) { /* segue */ }
    const regiao = (navigator.language || '').split('-')[1];
    if (regiao && MOEDA_POR_REGIAO[regiao.toUpperCase()]) return MOEDA_POR_REGIAO[regiao.toUpperCase()];
    return MOEDA_POR_IDIOMA[idioma] || 'USD';
  }
  function nonce() {
    const a = new Uint8Array(24); crypto.getRandomValues(a);
    return btoa(String.fromCharCode.apply(null, a)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  /* ---------- estado ---------- */
  let idioma = 'pt', moeda = 'BRL', valor = null, personalizado = false, cambio = null, cambioErro = false, config = null, ocupado = false;
  let t = (k) => k, pix = null; // injetados pela página
  const $ = (id) => document.getElementById(id);
  /* ?teste=1 na página: usa o checkout de teste da Stripe (nenhuma cobrança real) */
  const TESTE = /[?&]teste=1(&|$)/.test(location.search);

  async function carregaCambio() {
    if (cambio) return cambio;
    try {
      const r = await fetch(API + '/cambio', { cache: 'no-store' });
      if (!r.ok) throw new Error(String(r.status));
      cambio = await r.json(); cambioErro = false;
    } catch (e) { cambioErro = true; }
    return cambio;
  }
  async function carregaConfig() {
    if (config) return config;
    try { const r = await fetch(API + '/doacao/config' + (TESTE ? '?teste=1' : ''), { cache: 'no-store' }); if (r.ok) config = await r.json(); } catch (e) { config = null; }
    return config;
  }

  /* ---------- interface ---------- */
  function el(tag, attrs, filhos) {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) { if (k === 'text') n.textContent = v; else if (k === 'class') n.className = v; else if (k.startsWith('on')) n.addEventListener(k.slice(2), v); else if (v != null) n.setAttribute(k, v); }
    (filhos || []).forEach((f) => f && n.appendChild(typeof f === 'string' ? document.createTextNode(f) : f));
    return n;
  }
  function render() {
    const painel = $('supportPanel'); if (!painel) return;
    const m = MOEDAS[moeda];
    /* seletor de moeda */
    const sel = $('supportCurrency');
    if (sel) {
      sel.textContent = '';
      for (const k of Object.keys(MOEDAS)) sel.appendChild(el('option', { value: k, text: formata(0, k, idioma).replace(/[\d.,\s ]+/g, '').trim() + ' ' + ((NOMES[idioma] || NOMES.en)[k]) + ' (' + k + ')' }));
      sel.value = moeda;
    }
    /* valores sugeridos */
    const caixa = $('supportValues');
    caixa.textContent = '';
    m.sugeridos.forEach((v) => caixa.appendChild(el('button', { type: 'button', class: 'chip-valor' + (!personalizado && valor === v ? ' is-on' : ''), 'aria-pressed': String(!personalizado && valor === v), text: formata(v, moeda, idioma), onclick: () => { valor = v; personalizado = false; const c = $('supportCustom'); if (c) c.value = ''; render(); } })));
    caixa.appendChild(el('button', { type: 'button', class: 'chip-valor' + (personalizado ? ' is-on' : ''), 'aria-pressed': String(personalizado), text: t('supportOther'), onclick: () => { personalizado = true; valor = interpreta($('supportCustom').value, moeda, idioma); render(); setTimeout(() => $('supportCustom').focus(), 0); } }));
    /* campo personalizado */
    const campo = $('supportCustomBox');
    campo.hidden = !personalizado;
    const entrada = $('supportCustom');
    entrada.placeholder = formata(m.sugeridos[1], moeda, idioma);
    const [min, max] = LIMITES[moeda];
    const aviso = $('supportCustomHint');
    if (personalizado) {
      const bruto = entrada.value.trim();
      const v = interpreta(bruto, moeda, idioma);
      if (!bruto) aviso.textContent = t('supportCustomHelp').replace('{min}', formata(min, moeda, idioma)).replace('{max}', formata(max, moeda, idioma));
      else if (v === null) aviso.textContent = t('supportInvalid');
      else if (v < min) aviso.textContent = t('supportMin').replace('{min}', formata(min, moeda, idioma));
      else if (v > max) aviso.textContent = t('supportMax').replace('{max}', formata(max, moeda, idioma));
      else aviso.textContent = formataComCodigo(v, moeda, idioma);
      aviso.classList.toggle('erro', !!bruto && (v === null || v < min || v > max));
      valor = v !== null && v >= min && v <= max ? v : null;
    }
    /* cotação informativa */
    const cot = $('supportRate');
    if (cot) {
      if (moeda === 'BRL') cot.textContent = '';
      else if (cambio && cambio.taxas && valor) {
        const emBrl = converte(valor, moeda, 'BRL', cambio.taxas);
        const data = new Date(cambio.atualizadoEm);
        cot.textContent = t('supportApprox').replace('{valor}', formataComCodigo(emBrl, 'BRL', idioma)) + ' · ' + t('supportRateDate').replace('{data}', data.toLocaleString(locale(idioma, moeda), { dateStyle: 'short', timeStyle: 'short' })) + (cambio.velho ? ' · ' + t('supportRateOld') : '');
      } else if (cambioErro) cot.textContent = t('supportRateUnavailable');
      else cot.textContent = '';
    }
    /* formas de pagamento */
    const blocoPix = $('supportPix');
    const ehPix = moeda === 'BRL';
    blocoPix.hidden = !ehPix;
    if (ehPix && pix) pix(valor); // a página desenha o QR com o valor em reais (ou sem valor)
    const btnCartao = $('supportStripe');
    const podeCheckout = !!(config && config.checkout);
    btnCartao.hidden = false;
    btnCartao.disabled = ocupado || (podeCheckout && !valor);
    btnCartao.textContent = ocupado ? t('supportRedirecting') : t('supportStripe');
    const dica = $('supportCardHint');
    dica.textContent = podeCheckout ? t('supportCardExact') : t('supportCardHint');
    const metodos = $('supportMethods');
    if (metodos) metodos.textContent = t('supportMethods') + ' ' + (ehPix ? 'PIX · ' : '') + t('supportMethodCard');
    const resumo = $('supportSummary');
    if (resumo) resumo.textContent = (config && config.teste ? t('supportTestMode') + ' · ' : '') + (valor ? t('supportSummary').replace('{valor}', formataComCodigo(valor, moeda, idioma)) : '');
    const erro = $('supportError'); if (erro && !erro.dataset.fixo) erro.hidden = true;
  }

  async function iniciarCheckout(moedaCobranca, valorCobranca, confirmado) {
    const erro = $('supportError');
    erro.hidden = true; delete erro.dataset.fixo;
    if (!(config && config.checkout)) { window.open(config && config.paymentLink ? config.paymentLink : 'https://donate.stripe.com/aFaeVdgxw9mybwk2cB8bS09', '_blank', 'noopener'); return; }
    if (!valorCobranca) return;
    ocupado = true; render();
    try {
      const r = await fetch(API + '/doacao/sessao', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ moeda: moedaCobranca, valor: valorCobranca, idioma, nonce: nonce(), teste: TESTE }) });
      const j = await r.json().catch(() => ({}));
      if (r.ok && j.url && j.moeda === moedaCobranca && j.valor === valorCobranca) {
        /* guarda o que foi pedido para a página de obrigado conferir */
        try { sessionStorage.setItem('gifLocalDoacao', JSON.stringify({ moeda: moedaCobranca, valor: valorCobranca, idioma })); } catch (e) { /* segue */ }
        window.location.href = j.url; return;
      }
      if (r.status === 409 && j.error === 'MOEDA_NAO_ACEITA' && !confirmado) {
        /* a conta não cobra nesta moeda: oferece a moeda efetiva com o valor
           convertido, e só segue com confirmação explícita */
        await carregaCambio();
        const efetiva = j.moedaEfetiva || 'BRL';
        const convertido = cambio && cambio.taxas ? converte(valorCobranca, moedaCobranca, efetiva, cambio.taxas) : null;
        if (!convertido) { mostraErro(t('supportNoCurrencyNoRate').replace('{moeda}', moedaCobranca)); return; }
        mostraConfirmacao(efetiva, convertido, moedaCobranca, valorCobranca);
        return;
      }
      if (r.status === 503) { config.checkout = false; render(); mostraErro(t('supportCheckoutFallback')); return; }
      if (j.error === 'VALOR_INVALIDO') { mostraErro(t('supportMin').replace('{min}', formata(j.min, moedaCobranca, idioma)) + ' · ' + t('supportMax').replace('{max}', formata(j.max, moedaCobranca, idioma))); return; }
      mostraErro(t('supportCheckoutFail'));
    } catch (e) { mostraErro(t('supportCheckoutFail')); }
    finally { ocupado = false; render(); }
  }
  function mostraErro(msg) { const e = $('supportError'); e.textContent = msg; e.hidden = false; e.dataset.fixo = '1'; }
  function mostraConfirmacao(efetiva, convertido, pedidaMoeda, pedidoValor) {
    const box = $('supportConfirm');
    box.textContent = '';
    const data = cambio ? new Date(cambio.atualizadoEm).toLocaleString(locale(idioma, efetiva), { dateStyle: 'short', timeStyle: 'short' }) : '';
    box.appendChild(el('p', { text: t('supportEffective').replace('{pedida}', formataComCodigo(pedidoValor, pedidaMoeda, idioma)).replace('{efetiva}', formataComCodigo(convertido, efetiva, idioma)).replace('{data}', data) }));
    box.appendChild(el('div', { class: 'btn-row' }, [
      el('button', { type: 'button', class: 'btn', text: t('supportConfirmEffective').replace('{valor}', formataComCodigo(convertido, efetiva, idioma)), onclick: () => { box.hidden = true; iniciarCheckout(efetiva, convertido, true); } }),
      el('button', { type: 'button', class: 'btn sec-pay', text: t('supportCancel'), onclick: () => { box.hidden = true; } }),
    ]));
    box.hidden = false;
  }

  function liga(opts) {
    t = opts.t || t; pix = opts.pix || null; idioma = opts.idioma || 'pt';
    moeda = moedaInicial(idioma);
    valor = MOEDAS[moeda].sugeridos[1];
    const sel = $('supportCurrency');
    if (sel) sel.addEventListener('change', () => { moeda = MOEDAS[sel.value] ? sel.value : moeda; try { localStorage.setItem(CHAVE_MOEDA, moeda); } catch (e) { /* segue */ } valor = personalizado ? interpreta($('supportCustom').value, moeda, idioma) : MOEDAS[moeda].sugeridos[1]; render(); });
    const campo = $('supportCustom');
    if (campo) { campo.addEventListener('input', render); campo.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); if (valor) iniciarCheckout(moeda, valor, false); } }); }
    const btn = $('supportStripe');
    if (btn) btn.addEventListener('click', (e) => { e.preventDefault(); iniciarCheckout(moeda, valor, false); });
  }
  async function abre(opts) {
    if (opts && opts.idioma) idioma = opts.idioma;
    render();
    await Promise.all([carregaConfig(), carregaCambio()]);
    render();
  }
  function mudaIdioma(novo) { idioma = novo; let fixada = null; try { fixada = localStorage.getItem(CHAVE_MOEDA); } catch (e) { /* segue */ } if (!fixada) { moeda = moedaInicial(idioma); if (!personalizado) valor = MOEDAS[moeda].sugeridos[1]; } render(); }

  window.GifLocalDoacao = { liga, abre, render, mudaIdioma, formata, formataComCodigo, interpreta, converte, MOEDAS, LIMITES, separadores, get estado() { return { idioma, moeda, valor, personalizado }; } };
})();
