// Sobe a versão sempre que quiser forçar todo cliente instalado a descartar
// o cache antigo por inteiro (não só os arquivos que mudaram) — o activate
// abaixo já apaga qualquer cache com nome diferente deste.
const CACHE = 'aprosoja-custo-v6';
const FILES = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png', './Aprosoja-logo.png', './reference.json', './maquinas_marcas_modelos.json', './maquinas_anos_fabricacao.json'];

// Arquivos de dados/código que mudam com frequência (a cada ajuste do app ou
// a cada safra) — precisam ser buscados na rede primeiro, com o cache só
// como fallback offline. Ícones/logo/manifest praticamente não mudam, esses
// continuam cache-first (mais rápido, sem gasto de dados à toa).
const ARQUIVOS_NETWORK_FIRST = ['index.html', 'reference.json', 'maquinas_marcas_modelos.json', 'maquinas_anos_fabricacao.json'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  if (e.request.url.includes('api.github.com')) return;

  const ehNetworkFirst = e.request.mode === 'navigate' ||
    ARQUIVOS_NETWORK_FIRST.some(nome => e.request.url.includes(nome));

  if (ehNetworkFirst) {
    e.respondWith(
      fetch(e.request, { cache: 'no-cache' }).then(res => {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
        return res;
      }).catch(() =>
        caches.match(e.request).then(c => c || caches.match('./index.html'))
      )
    );
    return;
  }

  // Demais recursos (ícones, logo, manifest): cache-first
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(res => {
      if (res.ok && e.request.method === 'GET') {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
      }
      return res;
    }))
  );
});
