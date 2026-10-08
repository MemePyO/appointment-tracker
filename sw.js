"use strict";

const CACHE = "appointment-tracker-v1";
const CORE = [
  "./",
  "index.html",
  "style.css",
  "app.js",
  "manifest.webmanifest",
  "backend.py",
  "web_bridge.py",
  "icons/icon-192.png",
  "icons/icon-512.png",
  "icons/icon-maskable-512.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    await Promise.all(CORE.map((url) => cache.add(url).catch(() => {})));
    await self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);
  const allowed = url.origin === self.location.origin ||
                  url.hostname === "cdn.jsdelivr.net";
  if (!allowed) return;

  if (req.mode === "navigate") {
    event.respondWith((async () => {
      try {
        const res = await fetch(req);
        const copy = res.clone();
        const cache = await caches.open(CACHE);
        await cache.put("index.html", copy);
        return res;
      } catch (e) {
        const hit = await caches.match("index.html");
        if (hit) return hit;
        throw e;
      }
    })());
    return;
  }

  event.respondWith((async () => {
    const hit = await caches.match(req);
    if (hit) return hit;
    const res = await fetch(req);
    if (res && (res.ok || res.type === "opaque")) {
      const copy = res.clone();
      const cache = await caches.open(CACHE);
      await cache.put(req, copy);
    }
    return res;
  })());
});
