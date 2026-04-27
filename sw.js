// ============================================
// SERVICE WORKER — Dashboard Template PWA
// Versi cache: update angka jika ada file baru
// ============================================
const CACHE_NAME = 'dashboard-template-v3';

// File-file yang akan disimpan untuk offline
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './admin.html',
    './css/style.css',
    './js/app.js',
    './js/admin.js',
    './icons/icon-192.png',
    './icons/icon-512.png',
    './icons/icon-admin-192.png',
    './icons/icon-admin-512.png',
    './manifest.json',
    './manifest_admin.json'
];

// ---- INSTALL: Simpan semua aset ke cache ----
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            console.log('[SW] Menyimpan aset ke cache offline...');
            return cache.addAll(ASSETS_TO_CACHE);
        })
    );
    self.skipWaiting(); // Langsung aktif tanpa tunggu tab lama ditutup
});

// ---- ACTIVATE: Bersihkan cache lama ----
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames
                    .filter((name) => name !== CACHE_NAME)
                    .map((name) => {
                        console.log('[SW] Menghapus cache lama:', name);
                        return caches.delete(name);
                    })
            );
        })
    );
    self.clients.claim();
});

// ---- FETCH: Strategi "Network First, Cache Fallback" ----
// Prioritas: ambil dari internet (data segar), jika gagal pakai cache
self.addEventListener('fetch', (event) => {
    // Hanya cache request GET biasa (bukan API POST ke Google Apps Script)
    if (event.request.method !== 'GET') return;
    
    // API call ke Google Apps Script: selalu coba network, jangan cache
    if (event.request.url.includes('script.google.com')) return;

    event.respondWith(
        fetch(event.request)
            .then((networkResponse) => {
                // Sukses dari network: perbarui cache sambil jalan
                const cloned = networkResponse.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, cloned);
                });
                return networkResponse;
            })
            .catch(() => {
                // Gagal (offline): ambil dari cache
                return caches.match(event.request).then((cachedResponse) => {
                    if (cachedResponse) {
                        console.log('[SW] Offline mode: melayani dari cache —', event.request.url);
                        return cachedResponse;
                    }
                    // Jika tidak ada di cache sama sekali, kembalikan halaman utama
                    return caches.match('./index.html');
                });
            })
    );
});
