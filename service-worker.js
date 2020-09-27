/*
 * Service Worker for the SES (Simple Email Service) application.
 * The service worker is a basic application which intercepts network operations and either serves
 * the response for the operation from local cache, or forwards the request to the real internet.
 * In this way, the page can be displayed faster if a cached version is present (and the page can
 * potentially be updated later when the internet request returns, if one is sent).
 */

const CACHE_NAME = 'sesapp-cache-v1';
const URLS_TO_CACHE = [
    '.',
    'index.html',

    'config.json',
    'manifest.json',

    'gapi-worker.js'
];

/*
 * Install the PWA and service-worker. Create the cache object.
 * This is run once when the PWA is installed, and never run again.
 */
self.addEventListener('install', function(event) {
    event.waitUntil(
        // Register new cache and add all the requisite files to it
        caches.open(CACHE_NAME)
        .then(function(cache) {
            return cache.addAll(URLS_TO_CACHE);
        })
    );
    //take control immediately
    self.skipWaiting();
});

/*
 * Activate the app. Check for old cache versions and remove them.
 * This is run every time the PWA is loaded.
 */
self.addEventListener('activate', function(event) {
    event.waitUntil(
        caches.keys().then((keylist) => {
            // Delete old cache if one exists
            return Promise.all(keylist.map((key) => {
                if (key !== CACHE_NAME) {
                    console.log("[ServiceWorker] Removing old cache", key);
                    return caches.delete(key);
                }
            }));
        })
    );
    //take control immediately
    self.clients.claim();
});

/*
 * Intercept network fetch requests and optionally respond to those requests with information from
 * the cache, or if it's not present in the cache, send the request to the network.
 */
self.addEventListener('fetch', function(event) {
    if (event.request.url.includes("apis.google.com") ||
            event.request.url.includes("unpkg.com")) {
        return fetch(event.request.url, {mode: 'no-cors'});
    }
    console.log("[ServiceWorker] Fetch event:", event.request);
    event.respondWith(
        caches.match(event.request)
        .then(function(response) {
            return response || fetchAndCache(event.request);
        })
    );
});

function fetchAndCache(url) {
    return fetch(url)
    .then(function(response) {
        // Check if we received a valid response
        if (!response.ok) {
            throw Error(response.statusText);
        }
        return caches.open(CACHE_NAME)
        .then(function(cache) {
            cache.put(url, response.clone());
            return response;
        });
    })
    .catch(function(error) {
        console.log('Request failed:', url.url, error);
        // You could return a custom offline 404 page here
    });
}

