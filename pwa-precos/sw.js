const CACHE = 'aprosoja-precos-v22';
const FILES = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png', './Aprosoja-logo.png', './qr-app.png'];

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

  // Página: network-first, senão as atualizações do app nunca chegam a quem
  // já instalou (o cache antigo era servido para sempre — inclusive a versão
  // sem o leitor do link de configuração, o que impedia o token de ser salvo)
  if (e.request.mode === 'navigate' || e.request.url.includes('index.html')) {
    e.respondWith(
      // cache:'no-cache' pula o cache HTTP do GitHub Pages (10 min) — sem isso
      // a atualização demora a chegar mesmo com o network-first
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
