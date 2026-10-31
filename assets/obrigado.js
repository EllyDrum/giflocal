/* Página /obrigado: confirmação de doação (consulta real ao gateway) ou entrega
   de licença (fluxo antigo do PRO). Nada aqui afirma pagamento sem a resposta
   do servidor. */
(function () {
  'use strict';
  var API = 'https://api.giflocal.com';
  var params = new URLSearchParams(location.search);
  var sessionId = params.get('session_id');
  var ehDoacao = params.get('doacao') === '1';
  var TESTE = params.get('teste') === '1';

  var TXT = {
    pt: { carregando: 'Confirmando seu apoio...', carregandoDesc: 'Estamos consultando o gateway de pagamento. Isso leva só alguns segundos.', titulo: 'Obrigado pelo apoio!', detalhe: 'Valor confirmado pelo gateway de pagamento, na moeda cobrada.', desc: 'Sua contribuição ajuda a manter o GIF Local gratuito, sem anúncios e sem enviar arquivos para servidores.', voltar: 'Voltar ao GIF Local', teste: 'Modo de teste: nenhuma cobrança real', processandoT: 'Pagamento em processamento', processando: 'O gateway ainda está processando o pagamento. Alguns métodos levam mais tempo. Nenhum valor é dado como pago antes da confirmação.', abertoT: 'Pagamento não concluído', aberto: 'O pagamento ainda não foi concluído. Se você fechou o checkout, nada foi cobrado.', expiradoT: 'Sessão expirada', expirado: 'Esta sessão de pagamento expirou e nada foi cobrado.', erroT: 'Não foi possível confirmar', erro: 'Não conseguimos consultar o gateway agora. Se o pagamento foi feito, você receberá o recibo da Stripe por e-mail.', semId: 'Link incompleto: falta o identificador do pagamento.', tentar: 'Tentar de novo' },
    en: { carregando: 'Confirming your support...', carregandoDesc: 'We are checking with the payment gateway. This only takes a few seconds.', titulo: 'Thank you for your support!', detalhe: 'Amount confirmed by the payment gateway, in the charged currency.', desc: 'Your contribution helps keep GIF Local free, ad-free and without sending files to servers.', voltar: 'Back to GIF Local', teste: 'Test mode: no real charge', processandoT: 'Payment processing', processando: 'The gateway is still processing the payment. Some methods take longer. Nothing is shown as paid before confirmation.', abertoT: 'Payment not completed', aberto: 'The payment has not been completed. If you closed the checkout, nothing was charged.', expiradoT: 'Session expired', expirado: 'This payment session expired and nothing was charged.', erroT: 'Could not confirm', erro: 'We could not reach the gateway right now. If the payment went through, you will receive the Stripe receipt by email.', semId: 'Incomplete link: the payment identifier is missing.', tentar: 'Try again' },
    es: { carregando: 'Confirmando tu apoyo...', carregandoDesc: 'Estamos consultando la pasarela de pago. Solo tarda unos segundos.', titulo: '¡Gracias por tu apoyo!', detalhe: 'Importe confirmado por la pasarela de pago, en la moneda cobrada.', desc: 'Tu aportación ayuda a mantener GIF Local gratuito, sin anuncios y sin enviar archivos a servidores.', voltar: 'Volver a GIF Local', teste: 'Modo de prueba: ningún cargo real', processandoT: 'Pago en proceso', processando: 'La pasarela aún está procesando el pago. Algunos métodos tardan más. Nada se muestra como pagado antes de la confirmación.', abertoT: 'Pago no completado', aberto: 'El pago no se ha completado. Si cerraste el checkout, no se cobró nada.', expiradoT: 'Sesión caducada', expirado: 'Esta sesión de pago caducó y no se cobró nada.', erroT: 'No se pudo confirmar', erro: 'No pudimos consultar la pasarela ahora. Si el pago se realizó, recibirás el recibo de Stripe por correo.', semId: 'Enlace incompleto: falta el identificador del pago.', tentar: 'Intentar de nuevo' },
    fr: { carregando: 'Confirmation de votre soutien...', carregandoDesc: 'Nous interrogeons la passerelle de paiement. Cela ne prend que quelques secondes.', titulo: 'Merci pour votre soutien !', detalhe: 'Montant confirmé par la passerelle de paiement, dans la devise débitée.', desc: 'Votre contribution aide à garder GIF Local gratuit, sans publicité et sans envoi de fichiers vers des serveurs.', voltar: 'Retour à GIF Local', teste: 'Mode test : aucun débit réel', processandoT: 'Paiement en cours', processando: 'La passerelle traite encore le paiement. Certains moyens prennent plus de temps. Rien n’est affiché comme payé avant confirmation.', abertoT: 'Paiement non finalisé', aberto: 'Le paiement n’a pas été finalisé. Si vous avez fermé le checkout, rien n’a été débité.', expiradoT: 'Session expirée', expirado: 'Cette session de paiement a expiré et rien n’a été débité.', erroT: 'Confirmation impossible', erro: 'Impossible de joindre la passerelle pour le moment. Si le paiement a abouti, vous recevrez le reçu Stripe par e-mail.', semId: 'Lien incomplet : l’identifiant du paiement manque.', tentar: 'Réessayer' },
    de: { carregando: 'Ihre Unterstützung wird bestätigt...', carregandoDesc: 'Wir fragen beim Zahlungsanbieter nach. Das dauert nur ein paar Sekunden.', titulo: 'Danke für Ihre Unterstützung!', detalhe: 'Betrag vom Zahlungsanbieter bestätigt, in der belasteten Währung.', desc: 'Ihr Beitrag hilft, GIF Local kostenlos, werbefrei und ohne Datei-Uploads zu halten.', voltar: 'Zurück zu GIF Local', teste: 'Testmodus: keine echte Belastung', processandoT: 'Zahlung wird verarbeitet', processando: 'Der Zahlungsanbieter verarbeitet die Zahlung noch. Manche Methoden dauern länger. Vor der Bestätigung wird nichts als bezahlt angezeigt.', abertoT: 'Zahlung nicht abgeschlossen', aberto: 'Die Zahlung wurde nicht abgeschlossen. Wenn Sie den Checkout geschlossen haben, wurde nichts belastet.', expiradoT: 'Sitzung abgelaufen', expirado: 'Diese Zahlungssitzung ist abgelaufen, es wurde nichts belastet.', erroT: 'Bestätigung nicht möglich', erro: 'Der Zahlungsanbieter ist gerade nicht erreichbar. Wenn die Zahlung erfolgt ist, erhalten Sie die Stripe-Quittung per E-Mail.', semId: 'Unvollständiger Link: Die Zahlungskennung fehlt.', tentar: 'Erneut versuchen' },
    it: { carregando: 'Conferma del tuo supporto...', carregandoDesc: 'Stiamo verificando con il gateway di pagamento. Ci vogliono pochi secondi.', titulo: 'Grazie per il tuo supporto!', detalhe: 'Importo confermato dal gateway di pagamento, nella valuta addebitata.', desc: 'Il tuo contributo aiuta a mantenere GIF Local gratuito, senza pubblicità e senza inviare file ai server.', voltar: 'Torna a GIF Local', teste: 'Modalità test: nessun addebito reale', processandoT: 'Pagamento in elaborazione', processando: 'Il gateway sta ancora elaborando il pagamento. Alcuni metodi richiedono più tempo. Nulla viene mostrato come pagato prima della conferma.', abertoT: 'Pagamento non completato', aberto: 'Il pagamento non è stato completato. Se hai chiuso il checkout, non è stato addebitato nulla.', expiradoT: 'Sessione scaduta', expirado: 'Questa sessione di pagamento è scaduta e non è stato addebitato nulla.', erroT: 'Impossibile confermare', erro: 'Non riusciamo a contattare il gateway ora. Se il pagamento è andato a buon fine, riceverai la ricevuta Stripe via e-mail.', semId: 'Link incompleto: manca l’identificativo del pagamento.', tentar: 'Riprova' },
    ja: { carregando: 'ご支援を確認しています...', carregandoDesc: '決済ゲートウェイに確認中です。数秒で完了します。', titulo: 'ご支援ありがとうございます！', detalhe: '決済ゲートウェイが確認した、実際に請求された通貨での金額です。', desc: 'ご支援は、GIF Localを無料・広告なし・ファイルをサーバーに送らない形で続けるために使われます。', voltar: 'GIF Localに戻る', teste: 'テストモード: 実際の請求は発生しません', processandoT: '支払いを処理中', processando: '決済ゲートウェイが支払いを処理しています。方法によっては時間がかかります。確認が取れるまで支払い済みとは表示しません。', abertoT: '支払いは完了していません', aberto: '支払いはまだ完了していません。チェックアウトを閉じた場合、請求は発生していません。', expiradoT: 'セッションの有効期限切れ', expirado: 'この支払いセッションは期限切れとなり、請求は発生していません。', erroT: '確認できませんでした', erro: '現在、決済ゲートウェイに接続できません。支払いが完了している場合は、Stripeから領収書がメールで届きます。', semId: 'リンクが不完全です。支払いの識別子がありません。', tentar: 'もう一度試す' },
  };
  var LOCALES = { pt: 'pt-BR', en: 'en-US', es: 'es-ES', fr: 'fr-FR', de: 'de-DE', it: 'it-IT', ja: 'ja-JP' };
  var DEC0 = { JPY: 1, CLP: 1 };

  function idioma() {
    var lg = null;
    try { var d = JSON.parse(sessionStorage.getItem('gifLocalDoacao') || 'null'); if (d && d.idioma) lg = d.idioma; } catch (e) { /* segue */ }
    if (!lg) try { lg = localStorage.getItem('gifLocalLang'); } catch (e) { /* segue */ }
    if (!lg) lg = (navigator.language || 'pt').slice(0, 2).toLowerCase();
    return TXT[lg] ? lg : 'en';
  }
  var lang = idioma();
  var T = TXT[lang];
  document.documentElement.lang = LOCALES[lang] === 'pt-BR' ? 'pt-BR' : lang;

  function $(id) { return document.getElementById(id); }
  function show(id) {
    ['loadingState', 'donationState', 'readyState', 'errorState'].forEach(function (s) { $(s).classList.toggle('hidden', s !== id); });
  }
  function formata(minor, moeda) {
    var dec = DEC0[moeda] ? 0 : 2;
    var n = minor / Math.pow(10, dec);
    var o = { style: 'currency', currency: moeda, minimumFractionDigits: dec, maximumFractionDigits: dec };
    var s;
    try { s = new Intl.NumberFormat(LOCALES[lang], Object.assign({ currencyDisplay: 'narrowSymbol' }, o)).format(n); }
    catch (e) { s = new Intl.NumberFormat(LOCALES[lang], o).format(n); }
    return s.replace(/￥/g, '¥') + ' (' + moeda + ')';
  }

  /* ---------- doação ---------- */
  var tentativas = 0;
  function consultaDoacao() {
    tentativas++;
    fetch(API + '/doacao/status?session_id=' + encodeURIComponent(sessionId), { cache: 'no-store' })
      .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
      .then(function (r) {
        if (!r.ok) throw new Error('status');
        var j = r.j;
        if (j.situacao === 'pago') {
          document.title = T.titulo + ' | GIF Local';
          $('donationTitle').textContent = T.titulo;
          $('donationValue').textContent = formata(j.valor, j.moeda);
          $('donationDetail').textContent = T.detalhe;
          $('donationDesc').textContent = T.desc;
          $('donationBack').textContent = T.voltar;
          if (j.teste) { $('donationTest').textContent = T.teste; $('donationTest').classList.remove('hidden'); }
          try { sessionStorage.removeItem('gifLocalDoacao'); } catch (e) { /* segue */ }
          show('donationState');
          return;
        }
        if (j.situacao === 'processando' && tentativas < 10) { setTimeout(consultaDoacao, 3000); return; }
        var chave = j.situacao === 'processando' ? 'processando' : j.situacao === 'expirado' ? 'expirado' : 'aberto';
        erro(T[chave + 'T'], T[chave], j.situacao === 'processando');
      })
      .catch(function () {
        if (tentativas < 4) { setTimeout(consultaDoacao, 2500); return; }
        erro(T.erroT, T.erro, true);
      });
  }
  function erro(titulo, msg, podeTentar) {
    $('errorTitle').textContent = titulo;
    $('errorMsg').textContent = msg;
    $('retryBtn').textContent = T.tentar;
    $('retryBtn').classList.toggle('hidden', !podeTentar);
    $('errorBack').textContent = T.voltar;
    $('errorBack').classList.remove('hidden');
    show('errorState');
  }

  /* ---------- licença (fluxo antigo) ---------- */
  var attempts = 0;
  function fetchLicense() {
    if (!sessionId) { show('errorState'); $('errorMsg').textContent = 'Link incompleto: falta o identificador da compra.'; return; }
    attempts++;
    fetch(API + '/license-by-session?session_id=' + encodeURIComponent(sessionId))
      .then(function (res) { return res.json().then(function (data) { return { ok: res.ok, data: data }; }); })
      .then(function (r) {
        if (r.ok) {
          $('keyBox').textContent = r.data.license_key || r.data.licenseKey;
          show('readyState');
          try { localStorage.setItem('gifLocalPendingActivationKey', r.data.license_key || r.data.licenseKey); } catch (e) { /* segue */ }
        } else if (attempts < 6) setTimeout(fetchLicense, 3000);
        else show('errorState');
      })
      .catch(function () { if (attempts < 6) setTimeout(fetchLicense, 3000); else show('errorState'); });
  }

  if (ehDoacao) {
    document.title = T.titulo + ' | GIF Local';
    $('loadingTitle').textContent = T.carregando;
    $('loadingDesc').textContent = T.carregandoDesc;
    $('donationBack').href = TESTE ? '/?teste=1' : '/';
    $('errorBack').href = $('donationBack').href;
    if (!sessionId || !/^cs_(test|live)_[A-Za-z0-9]{10,200}$/.test(sessionId)) erro(T.erroT, T.semId, false);
    else consultaDoacao();
    $('retryBtn').addEventListener('click', function () { show('loadingState'); tentativas = 0; consultaDoacao(); });
  } else {
    fetchLicense();
    $('retryBtn').addEventListener('click', function () { show('loadingState'); attempts = 0; fetchLicense(); });
  }

  $('copyBtn').addEventListener('click', function () {
    var key = $('keyBox').textContent;
    var status = $('copyStatus');
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(key).then(function () { status.textContent = 'Copiado!'; status.className = 'status'; })
        .catch(function () { status.textContent = 'Não foi possível copiar. Selecione o texto manualmente.'; status.className = 'status error'; });
    } else { status.textContent = 'Selecione o texto acima manualmente para copiar.'; status.className = 'status error'; }
  });
})();
