/* ============================================================
   WT Paperless — Service Worker
   ทำให้แอปเปิดได้แม้ไม่มีเน็ต และติดตั้งลงหน้าจอโฮมได้
   หมายเหตุ: เบราว์เซอร์จะลงทะเบียนไฟล์นี้เฉพาะเมื่อเปิดผ่าน
             https:// หรือ http://localhost เท่านั้น
   ============================================================ */
const VERSION = "wt-paperless-v7";
const SHELL = [
  "./",
  "./index.html",                    /* ชื่อไฟล์บน GitHub Pages */
  "./WT-Paperless-WebApp.html",      /* ชื่อไฟล์บนเครื่องแม่ */
  "./WT-Paperless.webmanifest",
  "./WT-Paperless-Icons/icon-192.png",
  "./WT-Paperless-Icons/icon-512.png",
  "./WT-Paperless-Icons/icon-maskable-192.png",
  "./WT-Paperless-Icons/icon-maskable-512.png",
  "./WT-Paperless-Icons/apple-touch-icon.png",
  "./WT-Paperless-Icons/favicon-32.png"
];

self.addEventListener("install", e=>{
  e.waitUntil((async ()=>{
    const c = await caches.open(VERSION);
    /* เก็บทีละไฟล์ เพื่อไม่ให้ไฟล์เดียวที่หายทำให้ล้มทั้งชุด */
    await Promise.all(SHELL.map(u=> c.add(new Request(u, {cache:"reload"})).catch(()=>{})));
    self.skipWaiting();
  })());
});

self.addEventListener("activate", e=>{
  e.waitUntil((async ()=>{
    const keys = await caches.keys();
    await Promise.all(keys.filter(k=> k !== VERSION).map(k=> caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener("message", e=>{
  if(e.data === "skipWaiting") self.skipWaiting();
});

self.addEventListener("fetch", e=>{
  const req = e.request;
  if(req.method !== "GET") return;

  const url = new URL(req.url);
  if(url.origin !== self.location.origin) return;

  /* ── ข้อมูลสด: ห้ามแคชเด็ดขาด ── */
  if(url.pathname.startsWith("/api/")) return;

  /* ── ที่อยู่เครื่องแม่ล่าสุด: ต้องเอาของสดเสมอ ไม่งั้นมือถือจะจำที่อยู่เก่า ── */
  if(/api\.json$/.test(url.pathname)) return;

  /* ── รูปลายเซ็นย้อนหลัง (มีหลายหมื่นไฟล์): ผ่านไปตรง ๆ ── */
  if(/_Images\//.test(url.pathname)) return;

  /* ── หน้าเว็บ: เอาของใหม่ก่อน ถ้าเน็ตไม่มาค่อยใช้ของที่เก็บไว้ ── */
  const isDoc = req.mode === "navigate" || /\.html?$/.test(url.pathname) || url.pathname === "/";
  if(isDoc){
    e.respondWith((async ()=>{
      try{
        const fresh = await fetch(req);
        const c = await caches.open(VERSION);
        c.put(new URL("./", self.location).href, fresh.clone()).catch(()=>{});
        return fresh;
      }catch(err){
        const c = await caches.open(VERSION);
        return (await c.match(new URL("./", self.location).href)) ||
               (await c.match("./index.html")) || (await c.match("./WT-Paperless-WebApp.html")) ||
               new Response("<h1>ยังไม่ได้เชื่อมต่อ</h1><p>เปิดแอปครั้งแรกต้องต่อกับเครื่องแม่ก่อน</p>",
                            {headers:{"Content-Type":"text/html; charset=utf-8"}, status:503});
      }
    })());
    return;
  }

  /* ── ไฟล์อื่น (ไอคอน / manifest): ใช้ของที่เก็บไว้ก่อน ── */
  e.respondWith((async ()=>{
    const c = await caches.open(VERSION);
    const hit = await c.match(req);
    if(hit) return hit;
    try{
      const fresh = await fetch(req);
      if(fresh.ok) c.put(req, fresh.clone()).catch(()=>{});
      return fresh;
    }catch(err){
      return new Response("", {status:504});
    }
  })());
});
