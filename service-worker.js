// service-worker.js — AuditWorkpaper Pro Offline Engine
const CACHE_NAME = 'mus-template-v5';

const PRECACHE_ASSETS = [
  './',
  './index.html',
  './style.css',
  './script.js',
  './manifest.json',
  './DESIGN.md',
  './ico/icon-192.png',
  './ico/icon-512.png',
  './assets/Template_Input.xlsx',
  './assets/Template_Output.xlsx'
];

// External CDN Assets to cache for 100% offline capability
const CDN_HOSTS = [
  'cdn.jsdelivr.net',
  'cdnjs.cloudflare.com',
  'fonts.googleapis.com',
  'fonts.gstatic.com'
];

// Install Service Worker: Pre-cache App Shell & Assets
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('[Service Worker] Pre-caching local app shell...');
        return cache.addAll(PRECACHE_ASSETS);
      })
      .then(() => self.skipWaiting())
  );
});

// Activate Service Worker: Clean Up Old Caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keyList) => {
      return Promise.all(
        keyList.map((key) => {
          if (key !== CACHE_NAME) {
            console.log('[Service Worker] Removing deprecated cache:', key);
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch Strategy:
// 1. Navigation requests: Network First, falling back to cache if offline.
// 2. Static & CDN assets: Cache First with dynamic cache update.
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Network First for Navigation (HTML Page)
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request)
        .then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, responseToCache);
            });
          }
          return networkResponse;
        })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // Check if request is local or from supported CDNs
  const isLocal = url.origin === self.location.origin;
  const isCdn = CDN_HOSTS.some(host => url.hostname.includes(host));

  if (!isLocal && !isCdn) {
    return; // Pass through unknown third-party requests
  }

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }

      // Fetch from network and dynamically store in cache
      return fetch(event.request)
        .then((networkResponse) => {
          if (!networkResponse || (networkResponse.status !== 200 && networkResponse.type !== 'opaque')) {
            return networkResponse;
          }

          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });

          return networkResponse;
        })
        .catch(() => {
          return new Response('Offline Asset Unavailable', {
            status: 503,
            statusText: 'Service Unavailable'
          });
        });
    })
  );
});

// Listen for skip waiting messages
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});