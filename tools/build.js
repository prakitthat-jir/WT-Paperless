/* สร้าง/อัปเดต repo สำหรับขึ้น GitHub — คัดเฉพาะไฟล์โค้ด ไม่เอาข้อมูลไปด้วย
   รันซ้ำได้ทุกครั้งที่แก้แอปในโฟลเดอร์ออฟฟิศ */
const fs = require("fs"), path = require("path");

const OFFICE = "C:/Users/dissheet06/inter-group.co.th/PT IFC Rev.01 - PT 04.03.00 WT Paperless App ห้ามลบเด็ดขาด!!!";
const REPO   = process.env.WT_REPO || "C:/Users/dissheet06/wt-paperless";

/* ไฟล์โค้ดที่นำขึ้น GitHub ได้ (ไม่มีข้อมูลส่วนบุคคล) */
const COPY = [
  ["WT-Paperless-WebApp.html", "index.html"],
  ["WT-Paperless.webmanifest", "WT-Paperless.webmanifest"],
  ["WT-Paperless-SW.js",       "WT-Paperless-SW.js"],
  ["WT-Paperless-Server.js",   "server/WT-Paperless-Server.js"],
  ["START-WT-Paperless.bat",   "server/START-WT-Paperless.bat"],
  ["อ่านก่อนใช้งาน.txt",        "server/อ่านก่อนใช้งาน.txt"]
];
const ICONS = "WT-Paperless-Icons";

/* ─────────────────────────────────────────────────────────────
   ยามเฝ้าประตู — ห้ามไฟล์ที่มีข้อมูลส่วนบุคคลหลุดขึ้น GitHub
   ───────────────────────────────────────────────────────────── */
function personalProbes(){
  const p = path.join(OFFICE, "WT-Paperless-Base.json");
  if(!fs.existsSync(p)) return [];
  const b = JSON.parse(fs.readFileSync(p, "utf8"));
  const out = [];
  const push = (label, v)=>{ if(v && String(v).length >= 5) out.push([label, String(v)]); };
  (b.master.contractors || []).slice(0, 12).forEach((c,i)=>{
    push("ชื่อผู้รับเหมา#" + (i+1), c.name);
    push("ที่อยู่#" + (i+1), c.addr);
    push("โทร#" + (i+1), c.tel1);
  });
  (b.master.employees || []).slice(0, 12).forEach((e,i)=>{
    push("ชื่อพนักงาน#" + (i+1), typeof e === "string" ? e : e.name);
    if(e && e.tel) push("โทรพนักงาน#" + (i+1), e.tel);
  });
  (b.master.inspectors || []).forEach((x,i)=> push("ผู้ตรวจ#" + (i+1), x));
  (b.master.plates || []).slice(0, 20).forEach((x,i)=> push("ทะเบียน#" + (i+1), x));
  return out;
}
function scan(dir, probes){
  const hits = [];
  const walk = d =>{
    for(const name of fs.readdirSync(d)){
      if(name === ".git" || name === "node_modules") continue;
      const f = path.join(d, name);
      const st = fs.statSync(f);
      if(st.isDirectory()){ walk(f); continue; }
      if(/\.(png|jpg|jpeg|gif|ico|woff2?)$/i.test(name)) continue;
      const txt = fs.readFileSync(f, "utf8");
      probes.forEach(([label, v])=>{ if(txt.includes(v)) hits.push(path.relative(dir, f) + "  ←  " + label + ' "' + v.slice(0,40) + '"'); });
    }
  };
  walk(dir);
  return hits;
}

/* ─────────────────────────────────────────────────────────────
   ลงมือคัดลอก
   ───────────────────────────────────────────────────────────── */
function cp(from, to){
  fs.mkdirSync(path.dirname(to), {recursive:true});
  fs.copyFileSync(from, to);
}
fs.mkdirSync(REPO, {recursive:true});
let copied = 0;
COPY.forEach(([src, dst])=>{
  const from = path.join(OFFICE, src);
  if(!fs.existsSync(from)){ console.log("  ! ไม่พบ " + src); return; }
  cp(from, path.join(REPO, dst));
  copied++;
  console.log("  " + src.padEnd(30) + " →  " + dst);
});
fs.mkdirSync(path.join(REPO, ICONS), {recursive:true});
fs.readdirSync(path.join(OFFICE, ICONS)).forEach(f=>{
  cp(path.join(OFFICE, ICONS, f), path.join(REPO, ICONS, f)); copied++;
});
console.log("  " + ICONS.padEnd(30) + " →  " + ICONS + "/  (" + fs.readdirSync(path.join(REPO, ICONS)).length + " ไฟล์)");

/* GitHub Pages: ไม่ต้องประมวลผลด้วย Jekyll */
fs.writeFileSync(path.join(REPO, ".nojekyll"), "");

console.log("\n  คัดลอก " + copied + " ไฟล์ไปที่ " + REPO);

/* ─────────────────────────────────────────────────────────────
   ตรวจก่อนปล่อย
   ───────────────────────────────────────────────────────────── */
const probes = personalProbes();
console.log("\n  ตรวจข้อมูลส่วนบุคคลใน repo (" + probes.length + " รูปแบบ):");
const hits = scan(REPO, probes);
if(hits.length){
  hits.forEach(h=> console.log("   ⚠ " + h));
  console.log("\n   *** ห้าม push จนกว่าจะแก้ ***");
  process.exit(1);
}
console.log("   ✓ ไม่พบชื่อ / ที่อยู่ / เบอร์โทร / อีเมล / ทะเบียนรถ ในไฟล์ใด ๆ");

const total = (function size(d){ let n=0; for(const f of fs.readdirSync(d)){ if(f===".git")continue;
  const p=path.join(d,f), st=fs.statSync(p); n += st.isDirectory()? size(p) : st.size; } return n; })(REPO);
console.log("   ขนาด repo รวม " + (total/1024).toFixed(0) + " KB");
