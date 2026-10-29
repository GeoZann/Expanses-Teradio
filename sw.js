const CACHE = "tetradio-v6";
// Εκτός από τα δικά μας αρχεία, κρατάμε για offline μόνο βιβλιοθήκες και γραμματοσειρές.
// Οι κλήσεις του Firebase (googleapis.com) περνούν απευθείας, γιατί είναι ροές δεδομένων.
const CDN = ["www.gstatic.com", "cdn.jsdelivr.net", "fonts.googleapis.com", "fonts.gstatic.com"];
const ASSETS = ["./", "index.html", "import.js", "cloud.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "maskable-512.png", "apple-touch-icon.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE && k !== "share").map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

// Share target: Android στέλνει τα αρχεία με POST στο ./share. Το GitHub Pages δεν δέχεται POST,
// οπότε τα κρατάμε στην cache "share" και ανοίγουμε την εφαρμογή με ?share=1 για να τα διαβάσει.
async function share(req) {
  const form = await req.formData();
  await caches.delete("share");
  const c = await caches.open("share");
  const files = form.getAll("file");
  await Promise.all(files.map((f, i) =>
    c.put("shared-" + i, new Response(f, { headers: { "X-Name": encodeURIComponent(f.name || "αρχείο") } }))));
  return Response.redirect(new URL("./?share=1", self.registration.scope).href, 303);
}

// Network first, fall back to cache: updates arrive when online, app still opens offline.
self.addEventListener("fetch", (e) => {
  if (e.request.method === "POST" && new URL(e.request.url).pathname.endsWith("/share")) {
    e.respondWith(share(e.request));
    return;
  }
  const host = new URL(e.request.url).host;
  if (e.request.method !== "GET" || (host !== location.host && !CDN.includes(host))) return;
  e.respondWith(
    fetch(e.request)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request).then((r) => r || caches.match("./")))
  );
});
