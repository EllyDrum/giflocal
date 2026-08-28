/* Service worker do GIF Local.
   Faz cache do app inteiro no primeiro acesso, para funcionar de verdade
   offline depois — inclusive instalado como app (PWA) — sem nenhuma
   chamada de rede para processar fotos, vídeos ou gravações. */
const CACHE_NAME = 'gif-local-v4';
const ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './favicon-32.png',
  './apple-touch-icon.png',
  './brand-mark.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  /* NUNCA interceptar chamadas para outros domínios (a API de licenciamento
     e de geração com IA). Este service worker existe para deixar o APP
     disponível offline — respostas de API não são conteúdo estático e não
     devem ir para o cache.
     Sem esta linha, todo GET para a API passava por cache.put(), que rejeita
     em respostas de outra origem e derrubava a requisição inteira com
     "Failed to fetch". Sintoma real: consulta de cota da IA e a tela de
     "gerenciar dispositivos" simplesmente não funcionavam — enquanto POST
     (ativar licença) funcionava, porque sai na linha acima. */
  if (new URL(event.request.url).origin !== self.location.origin) return;

  /* Documento HTML (navegação): sempre tenta a rede primeiro, para que uma
     atualização do site apareça na próxima visita — só usa o cache se
     estiver offline. Isso evita ficar "preso" numa versão antiga. */
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          return response;
        })
        .catch(() => caches.match(event.request).then((cached) => cached || caches.match('./index.html')))
    );
    return;
  }

  /* Demais arquivos (ícones, manifest): cache primeiro, com atualização
     em segundo plano — são estáticos e raramente mudam. */
  event.respondWith(
    caches.match(event.request).then((cached) => {
      const fetchPromise = fetch(event.request)
        .then((response) => {
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => cached);
      return cached || fetchPromise;
    })
  );
});
