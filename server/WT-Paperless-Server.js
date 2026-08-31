/* ============================================================
   WT Paperless — ตัวช่วยบันทึกข้อมูลลงไฟล์ Excel ใหม่
   แผนกคลังสินค้าและจัดส่ง

   • ไม่แตะไฟล์ "Paperless of Warehouse & transportation.xlsx" เดิมเลย (อ่านก็ไม่อ่าน)
   • ข้อมูลที่กรอกใหม่ทั้งหมดจะถูกเขียนลงไฟล์ WT-Paperless-Data.xlsx
   • ใช้เฉพาะความสามารถที่มากับ Node.js ไม่ต้องติดตั้งอะไรเพิ่ม

   วิธีใช้: ดับเบิลคลิก "เปิดระบบ WT Paperless.bat"
   ============================================================ */

const http = require("http");
const fs   = require("fs");
const path = require("path");
const zlib = require("zlib");
const os   = require("os");
const crypto = require("crypto");
const { spawn } = require("child_process");

const HEADERS = {"การตรวจสอบรถขนส่ง ก่อน-หลัง":["การตรวจสอบรถขนส่ง ก่อน-หลัง ขึ้นสินค้า","ID","ปี","เดือน","วันที่","ทะเบียนรถ"," หมายเหตุ    \r\n1. รายการตรวจสอบก่อนขึ้นสินค้า และ 2.การป้องกันสินค้า หากรายการใดเป็น (NG=ไม่ผ่าน) จะต้องแจ้งให้คนรถดำเนินการแก้ไขให้เรียบร้อยก่อนทุกครั้ง จึงอนุญาตให้ขึ้นสินค้า และส่งสินค้าได้\r\n2. การตรวจสอบการป้องกันสินค้า ต้องเน้นความสมบูรณ์ของแผ่นพลาสติกที่กันการกระแทกมิให้มีร่องรอยการเสียหาย ฉีกขาด  (กระดาษเรียบไม่มีหัวน๊อตยื่นออกมา)","1. รายการตรวจสอบก่อนขึ้นสินค้า","1.1 ความสะอาดพื้นรถด้านใน\r\nมาตรฐานการตรวจสอบ: ต้องไม่มีเศษฝุ่น ขยะ และ เศษกระดาษ","1.2 ความสะอาดผนังรถด้านใน\r\nมาตรฐานการตรวจสอบ: ต้องไม่สกปรก มีคราบฝุ่น และ สิ่งแปลกปลอม","1.3 ความสะอาดของผ้าใบที่คลุมรถ\r\nมาตรฐานการตรวจสอบ: ต้องไม่สกปรก มีคราบฝุ่น และ สิ่งแปลกปลอม","1.4 ความพร้อมของรถขนส่ง (รอยรั่วน้ำมันเครื่อง)\r\nมาตรฐานการตรวจสอบ: ต้องไม่มีรอยรั่ว หยด ของน้ำมัน","1.5 ความพร้อมของเข็มขัดนิรภัย\r\nมาตรฐานการตรวจสอบ: มีเข็มขัดนิรภัยพร้อมใช้งาน","1.6 ความพร้อมของลมยางรถ\r\nมาตรฐานการตรวจสอบ: ยาง ลมยาง อยู่ในสภาพพร้อมใช้งาน (มองด้วยตา)","1.7 ความสมบูรณ์ของสินค้า (ไม่ชำรุด, เสียหาย)\r\nมาตรฐานการตรวจสอบ: สินค้าอยู่ในสภาพที่สมบูรณ์ ไม่ชำรุด ฉีกขาด","2. การป้องกันสินค้า","2.1 การปูวัสดุรองรับสินค้าบนพื้นรถ\r\nมาตรฐานการตรวจสอบ: พื้นรถต้องมีวัสดุปูรองพื้น","2.2 การปูวัสดุป้องกันสินค้าด้านผนังรถ\r\nมาตรฐานการตรวจสอบ: พนังรถต้องมีวัสดุป้องสินค้าชำรุด","2.3 การปูวัสดุปิดทับสินค้าด้านท้ายรถ\r\nมาตรฐานการตรวจสอบ: ท้ายรถรถต้องมีวัสดุป้องสินค้าชำรุด","2.4 ความพร้อมของผ้าใบ (รอยรั่ว ฉีกขาด)\r\nมาตรฐานการตรวจสอบ: ผ้าใบต้องอยู่ในสภาพพร้อมใช้ ไม่ชำรุด ฉีกขาด","2.5 ความสมบูรณ์ของพาเลท (ไม่ชำรุด, เสียหาย)\r\nมาตรฐานการตรวจสอบ: พาเลทอยู่ในสภาพพร้อมใช้งาน สะอาด ไม่ชำรุด","3. รายการตรวจสอบหลังขึ้นสินค้าเสร็จ","3.1 คลุมผ้าใบได้มิดชิด\r\nมาตรฐานการตรวจสอบ: ผ้าใบถูกคลุมสินค้าอย่างมิดชิด ไม่ก่อให้เกิดเสียหายต่อสินค้า","3.2 ปิดล็อคตู้มิดชิด\r\nมาตรฐานการตรวจสอบ: ฝาตู้ถูกปิดอย่างมิดชิด ไม่ก่อให้เกิดเสียหายต่อสินค้า","3.3 ไม่มีส่วนของสินค้าโผล่ออกมา/สูงเกินโดยไม่มีผ้าใบปกคลุม\r\nมาตรฐานการตรวจสอบ: สินค้าถูกผ้าใบคลุมมิดชิด ไม่ก่อให้เกิดความเสียหาย","3.4 รัดสายทางไกลแน่นหนา, ผูกมัดสินค้ากับตัวรถแน่นหนา\r\nมาตรฐานการตรวจสอบ: สายรัดเชือกถูกผูกแน่นหนา ไม่ก่อให้เกิดความเสี่ยง","ผู้ตรวจสอบ","เวลาที่ตรวจสอบ","ผู้ทวนสอบ","วันที่ทวนสอบ","FM-PB-007 Rev.03","วันที่บังคับใช้ : 09/03/63","ประเภทการจัดส่ง"],"การตรวจสุขลักษณะส่วนบุคคล":["การตรวจสุขลักษณะส่วนบุคคลคลังสินค้าและจัดส่ง","ID","ปี","เดือน","วันที่","เวลา","ชื่อพนักงาน","หมายเหตุ\r\n/ หมายถึง ถูกต้อง\r\nx หมายถึงไม่ถูกต้อง","ลักษณะการตรวจ\r\n1. การแต่งกาย ให้ถูกตามกฎระเบียบของบริษัท ห้ามใส่น้ำหอมที่มีกลิ่นฉุน\r\n2. เครื่องประดับ ห้ามสวมใส่ สร้อย แหวน ต่างหู นาฬิกา กำไร\r\n3. เล็บ ดูเรื่องความสะอาด เล็บยาว\r\n4. เส้นผม ห้ามใส่น้ำมันแต่งผม เจลแต่งผม","ไม่มีแผลที่มือ","การแต่งกาย","เครื่องประดับ","เล็บ","เส้นผม/หมวก","ear plug","การแก้ไข","ผู้ตรวจ","ผู้ทวนสอบ (การทวนสอบโดยหัวหน้าแผนกขึ้นไปเท่านั้น)","วันที่ทวนสอบ","FM-CG-029 Rev.00","วันที่บังคับใช้ : 01/11/63"],"การสำรวจวัสดุที่ทำด้วยแก้ว":["การสำรวจวัสดุที่ทำด้วยแก้ว และวัสดุที่แตกหักได้ 3 (คลังสินค้าและจัดส่ง)","No.","ปี","เดือน","วันที่","หมายเหตุ: \r\n1. / = ไม่พบสิ่งผิดปกติ\r\n2. GI = พบสิ่งผิดปกติ\r\n***ทวนสอบโดยหัวหน้าแผนกขึ้นไปเท่านั้น***","1. หลอดไฟกลม","1.1 คลังสินค้าและคลังสำเร็จรูป (G12 - G31)","2. หลอดไฟแสงสว่าง","2.1 คลังสินค้าและสำเร็จรูป","2.1.1 หลอดไฟให้แสงสว่าง","2.1.2 หลอดไฟดักแมลง ","2.2 ISOWA","2.2.1 หลอดไฟให้แสงสว่าง","3. คอมพิวเตอร์, ตู้ Control","3.1 คลังสินค้าและสำเร็จรูป","3.1.1 คอมพิวเตอร์ ","3.1.2 ตู้ Control","3.2 ISOWA","3.2.1 คอมพิวเตอร์ ","3.2.2 ตู้ Control","4. กล้องวงจรปิด","4.1 คลังสินค้าและสำเร็จรูป","5. หลอดไฟฉุกเฉิน และป้ายฉุกเฉิน","5.1 คลังสินค้าและสำเร็จรูป","6. นาฬิกา","6.1 เครื่อง ISOWA","7. เครื่องชั่งน้ำหนัก","7.1 เครื่อง ISOWA","8. กระจก, แก้วที่แตกหักได้","8.1 ห้องเตรียมบล็อกและสีพิมพ์","8.2 เครื่อง ISOWA","8.3 ห้องจัดส่ง","8.4 บานเกร็ดกระจกหลังศาล","8.5 ห้องผสมกาว","8.6 ห้องเก็บน้ำมันหล่อลื่น","9. ตู้ดับเพลิง","9.1 คลังสินค้าและสำเร็จรูป","การแก้ไข","รายงานโดย","ทวนสอบโดย หัวหน้าแผนกคลังสินค้าและจัดส่ง","FM-CG-027 Rev.01","วันที่บังคับใช้: 01/10/65"],"ใบตรวจสอบรถขนส่ง":["ใบตรวจสอบสอบสภาพรถขนส่ง","No.","รหัสผู้รับเหมา","YEAR","Month","Date","ชื่อผู้ให้บริการขนส่ง","ทะเบียนรถ","ประเภทการจัดส่ง","ข้อกำหนดในการตรวจสอบ (ด้านคุณภาพสิ่งแวดล้อมสุขลักษณะอาชีวอนามัยและความปลอดภัย)","1. สภาพรถ/สุขลักษณะ","1.1 ความสะอาดทั่วไปภายในและภายนอก แผ่นปูพื้น, แผ่นบุข้าง, เชือกมัดและผ้าใบ","1.2 ร่องรอยสัตว์พาหะ มด, หนู, แมลงสาบ, นกและแมลงบิน","1.3 ไฟส่องสว่าง หน้า, หลัง, เบรก, และไฟเลี้ยว","1.4 ใบปัดน้ำฝนใช้งานได้","1.5 ประเภททะเบียนรถ ต้องเป็นป้ายเหลือง","1.6 ความพร้อมของเข็มขัดนิรภัย","1.7 ความพร้อมของเบรค","1.8 ความพร้อมของยาง, สภาพยาง","2. ด้านชีวอนามัยและความปลอดภัย","2.1 ใบอนุญาตขับรถ ต้องมีติดตัว และไม่หมดอายุ","2.2 พรบ. คุ้มครองผู้ประสบภัย ต้องไม่หมดอายุ","2.3 หมอนไม้ล้อ กันรถไหล ต้องมีไม่ต่ำกว่า 2 ชิ้น","2.4 การแต่งกายของคนขับรถถูกระเบียบ","2.5 มีถุงมือกรณีกรณีขนส่งกล่องบรรจุอาหาร พร้อมและพอเพียงกับพนักงานประจำรถ","2.6 ความพร้อมของอุปกรณ์ป้องกัน หมวก รองเท้า","3. ด้านสิ่งแวดล้อม","3.1 ร่องรอยการรั่วไหลของน้ำมันหล่อลื่น","3.2 สภาพการปล่อยควันของท่อไอเสีย","คะแนนเต็ม","คะแนนรวม %","ผลการประเมิน","การดำเนินงาน","เหตุผล เฉพาะที่ต้องปรับปรุง","ผู้ตรวจสอบ หผ.จัดส่งและคลังสินค้า","วัน/เดือน/ปี ที่ตรวจสอบ","ผู้อนุมัติ ผจฝ.ส่งเสริมการผลิต","วัน/เดือน/ปี ที่อนุมัติ","FM-TB-012 Rev.00","วันที่บังคับใช้: 07/11/60"],"ประเมินการบริการขนส่ง":["แบบประเมินผู้รับเหมาในการบริการขนส่ง","No.","YEAR","Month","Date","รหัสผู้รับเหมา","ชื่อผู้รับเหมา","ประเภทรถขนส่ง","ประเภทการจัดส่ง","ครั้งที่ประเมิน","หัวข้อการประเมิน","1. คุณภาพ 60 คะแนน","1.1 ไม่มีการร้องเรียนจากทางลูกค้าเรื่องคุณภาพที่เกิดจากการบริการ และการส่งมอบ คะแนนเต็ม 40 คะแนน","1.2 มีการส่งมอบเอกสารภูกต้อง ลูกค้าไม่มีการร้องเรียน เช่น COA, เอกสารของแผนกจัดส่ง คะแนนเต็ม 10 คะแนน","1.3 สภาพของรถมีความปลอดภัย ไม่มีความเสี่ยงต่อการเกิดอุบัติเหตุต่อผลิตภัณฑ์ คะแนนเต็ม 10 คะแนน","2. การส่งมอบ 30 คะแนน","2.1 สินค้าส่งถึงลูกค้าตรงตามกำหนดเวลาไม่ติดปัญหา คะแนนเต็ม 10 คะแนน","2.2 สินค้าส่งครบตามจำนวนที่ตกลงกันไว้ คะแนนเต็ม 10 คะแนน","2.3 สามารถปฏิบัติตามกฏระเบียบของบริษัทฯ ได้ถูกต้อง คะแนนเต็ม 5 คะแนน","2.4 สามารถปฏิบัติตามกฏระเบียบของล฿กค้าไม่มีการร้องเรียน คะแนนเต็ม 5 คะแนน","3. การบริการ 10 คะแนน","3.1 มีการติดตามงาน และการบริการที่รวดเร็วตอบสนองความต้องการของลูกค้าได้ดี คะแนนเต็ม 5 คะแนน","3.2 สามารถแก้ไขปัญหาให้กับทางลูกค้าในกรณีการส่งสินค้าเข้าไปทดแทนได้รวดเร็ว คะแนนเต็ม 5 คะแนน","เกณฑ์การประเมิน","คะแนนเกรด","ผลเกรด","ระดับ","ผลการพิจารณา","ข้อเสนอแนะในการปรับปรุง","ผู้ประเมิน","วันที่ประเมิน","FM-QMR-021 Rev.01","วันที่บังคับใช้: 30/10/62"],"พนักงาน":["ลำดับ","ชื่อ - นามสกุล","ชื่อเล่น","ตำแหน่ง","เบอร์โทรศัพท์"],"ผู้รับเหมา":["รายละเอียดผู้รับเหมาในการบริการขนส่ง  (AVL)","ลำดับ","รหัสผู้รับเหมา","ชื่อผู้รับเหมา","ที่อยู่","ชื่อผู้ติดต่อ","ทะเบียนรถ","โทรศัพท์ 1","โทรศัพท์ 2","ลายเซ็นต์","ประเภทการจัดส่ง","FM-QMR-020 Rev. 00","วันที่บังคับใช้ : 03/01/56"],"ประเภทการจัดส่ง":["ประเภทการจัดส่ง"]};
const COLMAP  = {"การตรวจสอบรถขนส่ง ก่อน-หลัง":["","id","year","month","date","f:plate","","","i:1.1","i:1.2","i:1.3","i:1.4","i:1.5","i:1.6","i:1.7","","i:2.1","i:2.2","i:2.3","i:2.4","i:2.5","","i:3.1","i:3.2","i:3.3","i:3.4","f:inspector","time","f:verifier","f:verifyDate","","","f:delivery"],"การตรวจสุขลักษณะส่วนบุคคล":["","id","year","month","date","time","f:employee","","","i:ไม่มีแผลที่มือ","i:การแต่งกาย","i:เครื่องประดับ","i:เล็บ","i:เส้นผม/หมวก","i:ear plug","f:fix","f:inspector","f:verifier","f:verifyDate","",""],"การสำรวจวัสดุที่ทำด้วยแก้ว":["","no","year","month","date","","","i:1.1","","","i:2.1.1","i:2.1.2","","i:2.2.1","","","i:3.1.1","i:3.1.2","","i:3.2.1","i:3.2.2","","i:4.1","","i:5.1","","i:6.1","","i:7.1","","i:8.1","i:8.2","i:8.3","i:8.4","i:8.5","i:8.6","","i:9.1","f:fix","f:inspector","f:verifier","",""],"ใบตรวจสอบรถขนส่ง":["","no","f:contractorCode","year","month","date","f:contractor","f:plate","f:delivery","","","i:1.1","i:1.2","i:1.3","i:1.4","i:1.5","i:1.6","i:1.7","i:1.8","","i:2.1","i:2.2","i:2.3","i:2.4","i:2.5","i:2.6","","i:3.1","i:3.2","c:max","c:pct","c:result","c:action","f:reason","f:inspector","date","f:approver","f:approveDate","",""],"ประเมินการบริการขนส่ง":["","no","year","month","date","f:contractorCode","f:contractor","f:truckType","f:delivery","f:round","","","i:1.1","i:1.2","i:1.3","","i:2.1","i:2.2","i:2.3","i:2.4","","i:3.1","i:3.2","","c:sum","c:grade","c:level","c:gradeResult","f:suggest","f:inspector","date","",""]};
/* ข้อมูลจริงอยู่ในไฟล์ WT-Paperless-Base.json ข้าง ๆ ไฟล์นี้ (ไม่ขึ้น git)
   ไฟล์เซิร์ฟเวอร์จึงไม่มีข้อมูลส่วนบุคคลใด ๆ */
const MASTER  = {"contractors":[],"employees":[],"plates":[],"inspectors":[],"deliveryTypes":["BOX","SHEET"],"truckTypes":[]};

const ROOT      = __dirname;
const DATA_JSON = path.join(ROOT, "WT-Paperless-Data.json");
const DATA_XLSX = path.join(ROOT, "WT-Paperless-Data.xlsx");
const IMG_DIR   = path.join(ROOT, "WT-Paperless-Images");
const BACKUP    = path.join(ROOT, "WT-Paperless-Backup");
const APP_FILE  = "WT-Paperless-WebApp.html";
const PORT      = Number(process.env.WT_PORT || 8080);
const KEEP_BACKUPS = 40;
const BASE_JSON   = path.join(ROOT, "WT-Paperless-Base.json");     /* ข้อมูลตั้งต้น + ประวัติย้อนหลัง */
const CONFIG_JSON = path.join(ROOT, "WT-Paperless-Config.json");   /* รหัสเข้าใช้ — ห้ามขึ้น git */

/* ============================================================
   รหัสเข้าใช้งาน (ใช้เมื่อเปิดจากนอกออฟฟิศผ่านอุโมงค์)
   ============================================================ */
let CFG = null;
function loadConfig(){
  try{ if(fs.existsSync(CONFIG_JSON)) CFG = JSON.parse(fs.readFileSync(CONFIG_JSON,"utf8")); }catch(e){}
  if(!CFG || typeof CFG !== "object") CFG = {};
  let changed = false;
  if(!CFG.accessCode){                                  /* ตั้งรหัส 6 หลักให้อัตโนมัติครั้งแรก */
    CFG.accessCode = String(crypto.randomInt(100000, 1000000));
    changed = true;
  }
  if(!CFG.secret){ CFG.secret = crypto.randomBytes(32).toString("hex"); changed = true; }
  if(CFG.requireCodeOnLan === undefined) CFG.requireCodeOnLan = false;  /* ในออฟฟิศไม่ต้องใส่รหัส */
  if(changed) fs.writeFileSync(CONFIG_JSON, JSON.stringify(CFG, null, 2), "utf8");
  return CFG;
}
/* โทเคนแบบเซ็นชื่อ — ไม่ต้องเก็บ session ในหน่วยความจำ รีสตาร์ตแล้วยังใช้ต่อได้ */
function makeToken(days){
  const exp = Date.now() + (days || 30) * 86400000;
  const sig = crypto.createHmac("sha256", CFG.secret).update(String(exp)).digest("base64url");
  return exp + "." + sig;
}
function validToken(tok){
  if(typeof tok !== "string") return false;
  const [expStr, sig] = tok.split(".");
  const exp = Number(expStr);
  if(!exp || !sig || Date.now() > exp) return false;
  const want = crypto.createHmac("sha256", CFG.secret).update(String(exp)).digest("base64url");
  const a = Buffer.from(sig), b = Buffer.from(want);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}
/* กันเดารหัส */
const LOGIN_FAILS = new Map();
function loginBlocked(ip){
  const f = LOGIN_FAILS.get(ip);
  if(!f) return 0;
  if(Date.now() > f.until){ LOGIN_FAILS.delete(ip); return 0; }
  return f.count >= 5 ? Math.ceil((f.until - Date.now())/1000) : 0;
}
function noteLoginFail(ip){
  const f = LOGIN_FAILS.get(ip) || {count:0, until:0};
  f.count++; f.until = Date.now() + 5*60000;
  LOGIN_JAIL_TRIM();
  LOGIN_FAILS.set(ip, f);
}
function LOGIN_JAIL_TRIM(){ if(LOGIN_FAILS.size > 500) LOGIN_FAILS.clear(); }
function clientIP(req){
  return (req.socket && req.socket.remoteAddress) || "?";
}
function isLocalReq(req){
  const ip = clientIP(req);
  return ip === "127.0.0.1" || ip === "::1" || ip === "::ffff:127.0.0.1";
}

/* ============================================================
   ข้อมูลตั้งต้น (ข้อมูลหลัก + ประวัติย้อนหลัง) — ส่งให้แอปตอนเปิด
   ============================================================ */
let BASE = null, BASE_GZ = null, BASE_TAG = "";
function loadBase(){
  try{
    if(!fs.existsSync(BASE_JSON)) return;
    BASE = JSON.parse(fs.readFileSync(BASE_JSON, "utf8"));
    if(BASE && BASE.master) Object.assign(MASTER, BASE.master);   /* ข้อมูลตั้งต้นมาจากไฟล์ ไม่ได้ฝังในโค้ด */
  }catch(e){ console.log("  ! อ่าน WT-Paperless-Base.json ไม่ได้:", e.message); }
}
function bootstrapPayload(){
  const master = Object.assign({}, (BASE && BASE.master) || {});
  const eff = M();                                   /* ทับด้วยข้อมูลหลักที่ผู้ใช้แก้ไว้ */
  master.contractors = eff.contractors;
  master.employees   = eff.employees;
  if(DB.master && Array.isArray(DB.master.plates)) master.plates = DB.master.plates;
  return {ok:true, master, history:(BASE && BASE.history) || {}};
}
function bootstrapGz(){
  const tag = String(DB.savedAt || 0) + ":" + ((DB.master && DB.master.employees) || []).length +
              ":" + ((DB.master && DB.master.contractors) || []).length;
  if(BASE_GZ && BASE_TAG === tag) return BASE_GZ;
  BASE_GZ = zlib.gzipSync(Buffer.from(JSON.stringify(bootstrapPayload()), "utf8"), {level:6});
  BASE_TAG = tag;
  return BASE_GZ;
}

const SHEET = {
  vehicle:    "การตรวจสอบรถขนส่ง ก่อน-หลัง",
  hygiene:    "การตรวจสุขลักษณะส่วนบุคคล",
  glass:      "การสำรวจวัสดุที่ทำด้วยแก้ว",
  truckcheck: "ใบตรวจสอบรถขนส่ง",
  evaluate:   "ประเมินการบริการขนส่ง"
};
const FORM_ORDER = ["vehicle","hygiene","glass","truckcheck","evaluate"];

/* เกณฑ์คะแนน (ตรงกับที่ใช้ในเว็บแอป) */
const SCORE = {
  truckcheck: {items:16, perItem:3, passPct:90,
    pass:"ผ่าน", fail:"ต้องปรับปรุง",
    actionPass:"รับเป็นผู้ส่งมอบต่อไป",
    actionFail:"รับเป็นผู้ส่งมอบต่อไป และแจ้งให้ทำการแก้ไขปรับปรุง"},
  evaluate: {max:{"1.1":40,"1.2":10,"1.3":10,"2.1":10,"2.2":10,"2.3":5,"2.4":5,"3.1":5,"3.2":5},
    grades:[{min:98,g:"A",lv:"ดีมาก",res:"ต้องมีการรักษามาตรฐานดังกล่าวไว้ จะถูกเลือกเป็นตัวหลักในการเลือกเข้าให้บริการแต่ละครั้ง"},
            {min:90,g:"B",lv:"ดี",res:"ต้องมีการควบคุมให้เกิดการปรับปรุงให้เข้าสู่ระดับเกรด A"},
            {min:85,g:"C",lv:"พอใช้",res:"ต้องแจ้งให้ปรับปรุงการให้บริการ และติดตามผลในการประเมินครั้งถัดไป"},
            {min:80,g:"D",lv:"ปรับปรุง",res:"ต้องแจ้งให้แก้ไขปรับปรุงเป็นลายลักษณ์อักษร และลดปริมาณงานที่มอบหมายจนกว่าจะดีขึ้น"},
            {min:0, g:"F",lv:"ไม่ผ่าน",res:"ไม่ผ่านเกณฑ์ — ต้องพิจารณาทบทวนการใช้บริการ และระงับการมอบหมายงานจนกว่าจะแก้ไขแล้วเสร็จ"}]}
};

/* ============================================================
   ฐานข้อมูล (ไฟล์ JSON คือตัวจริง, .xlsx สร้างใหม่ทุกครั้งที่บันทึก)
   ============================================================ */
let DB = {records:{}, master:null, savedAt:null};
/* ข้อมูลหลักที่ใช้จริง = ที่ผู้ใช้แก้ไว้ (ถ้ามี) มิฉะนั้นใช้ชุดตั้งต้น */
const M = () => ({
  contractors: (DB.master && Array.isArray(DB.master.contractors)) ? DB.master.contractors : (MASTER.contractors||[]),
  employees:   normEmpList((DB.master && Array.isArray(DB.master.employees)) ? DB.master.employees : (MASTER.employees||[]))
});
/* พนักงานเก็บเป็น {name,nick,pos,tel} — รองรับข้อมูลเก่าที่เป็นข้อความล้วน */
const ROSTER_VER = 2;
function normEmp(v){
  if(typeof v === "string") return {name:v.trim(), nick:"", pos:"", tel:""};
  return {name:String((v&&v.name)||"").trim(), nick:String((v&&v.nick)||"").trim(),
          pos:String((v&&v.pos)||"").trim(),  tel:String((v&&v.tel)||"").trim()};
}
function normEmpList(a){ return (Array.isArray(a)?a:[]).map(normEmp).filter(e=> e.name); }
/* ครั้งแรกหลังอัปเดต: รวมผังรายชื่อชุดใหม่เข้ากับที่บันทึกไว้ (เติมคนที่ยังไม่มี + เติมชื่อเล่น/ตำแหน่ง/เบอร์) */
function migrateRoster(){
  if(DB.rosterVer === ROSTER_VER) return false;
  const seed  = normEmpList(MASTER.employees);
  const saved = normEmpList(DB.master && DB.master.employees);
  /* เทียบชื่อโดยตัดคำนำหน้าและช่องว่างออก เพื่อไม่ให้ “นายสมชาย” กับ “นาย สมชาย” กลายเป็นคนละคน */
  const key = s2 => String(s2||"").replace(/s+/g,"").replace(/^(นางสาว|นาง|นาย|น.ส.|นส.)/,"");
  const out = seed.slice();                                   /* ยึดลำดับตามผังรายชื่อ */
  const seen = new Set(out.map(e=> key(e.name)));
  saved.forEach(e=>{ if(!seen.has(key(e.name))){ out.push(e); seen.add(key(e.name)); } });
  DB.master = Object.assign({}, DB.master, {employees: out});
  DB.rosterVer = ROSTER_VER;
  console.log("  ปรับผังรายชื่อพนักงานเป็นชุดใหม่ — รวม " + out.length + " คน");
  return true;
}
function loadDB(){
  try{
    if(fs.existsSync(DATA_JSON)) DB = JSON.parse(fs.readFileSync(DATA_JSON,"utf8"));
  }catch(e){ console.log("  ! อ่านไฟล์ข้อมูลเดิมไม่ได้:", e.message); }
  FORM_ORDER.forEach(f=>{ if(!Array.isArray(DB.records[f])) DB.records[f] = []; });
  if(DB.master && Array.isArray(DB.master.employees)) DB.master.employees = normEmpList(DB.master.employees);
  if(migrateRoster()) persist();
}
function backup(){
  if(!fs.existsSync(DATA_JSON)) return;
  try{
    if(!fs.existsSync(BACKUP)) fs.mkdirSync(BACKUP);
    const st = new Date();
    const name = "data-" + st.getFullYear() + pad(st.getMonth()+1) + pad(st.getDate())
               + "-" + pad(st.getHours()) + pad(st.getMinutes()) + pad(st.getSeconds()) + ".json";
    fs.copyFileSync(DATA_JSON, path.join(BACKUP, name));
    const old = fs.readdirSync(BACKUP).filter(f=>f.startsWith("data-")).sort();
    while(old.length > KEEP_BACKUPS) fs.unlinkSync(path.join(BACKUP, old.shift()));
  }catch(e){ console.log("  ! สำรองข้อมูลไม่สำเร็จ:", e.message); }
}
let writing = false, pending = false;
function persist(){
  if(writing){ pending = true; return; }
  writing = true;
  try{
    backup();
    DB.savedAt = Date.now();
    fs.writeFileSync(DATA_JSON + ".tmp", JSON.stringify(DB), "utf8");
    fs.renameSync(DATA_JSON + ".tmp", DATA_JSON);
    const buf = buildWorkbook();
    fs.writeFileSync(DATA_XLSX + ".tmp", buf);
    try{ fs.renameSync(DATA_XLSX + ".tmp", DATA_XLSX); }
    catch(e){
      /* ไฟล์ถูกเปิดค้างใน Excel */
      fs.unlinkSync(DATA_XLSX + ".tmp");
      throw new Error("เขียนไฟล์ Excel ไม่ได้ — กรุณาปิดไฟล์ WT-Paperless-Data.xlsx ใน Excel ก่อน (ข้อมูลถูกเก็บไว้แล้ว จะเขียนลง Excel ให้อัตโนมัติในการบันทึกครั้งถัดไป)");
    }
  } finally {
    writing = false;
    if(pending){ pending = false; persist(); }
  }
}
const pad = n => String(n).padStart(2,"0");

/* ---------- ลายเซ็น: แปลง data URL เป็นไฟล์ PNG ---------- */
function saveSignatures(rec){
  ["verifier","inspector","approver"].forEach(k=>{
    const v = rec[k];
    if(typeof v === "string" && v.startsWith("data:image/")){
      try{
        if(!fs.existsSync(IMG_DIR)) fs.mkdirSync(IMG_DIR);
        const b64 = v.split(",")[1];
        const t = new Date();
        const file = rec.id + "." + k + "." + pad(t.getHours())+pad(t.getMinutes())+pad(t.getSeconds()) + ".png";
        fs.writeFileSync(path.join(IMG_DIR, file), Buffer.from(b64, "base64"));
        rec[k] = "WT-Paperless-Images/" + file;
      }catch(e){ console.log("  ! บันทึกลายเซ็นไม่สำเร็จ:", e.message); }
    }
  });
  return rec;
}

/* ============================================================
   คำนวณคะแนน
   ============================================================ */
function calc(fid, r){
  const it = r.items || {};
  if(fid === "truckcheck"){
    const c = SCORE.truckcheck;
    const sum = Object.keys(it).reduce((a,k)=> a + (Number(it[k])||0), 0);
    const max = c.items * c.perItem;
    const pct = Math.round(sum/max*10000)/100;
    const ok  = pct >= c.passPct;
    return {sum, max, pct, result: ok?c.pass:c.fail, action: ok?c.actionPass:c.actionFail};
  }
  if(fid === "evaluate"){
    const c = SCORE.evaluate;
    const sum = Object.keys(c.max).reduce((a,k)=> a + (Number(it[k])||0), 0);
    const g = c.grades.find(x=> sum >= x.min) || c.grades[c.grades.length-1];
    return {sum, max:100, grade:g.g, level:g.lv, gradeResult:g.res};
  }
  return {};
}

/* ============================================================
   ตัวเขียนไฟล์ .xlsx (ZIP + XML) — ไม่ใช้ไลบรารีภายนอก
   ============================================================ */
const CRC_TABLE = (()=>{ const t=new Array(256);
  for(let n=0;n<256;n++){ let c=n; for(let k=0;k<8;k++) c = (c&1)? (0xEDB88320 ^ (c>>>1)) : (c>>>1); t[n]=c>>>0; }
  return t; })();
function crc32(buf){
  let c = 0xFFFFFFFF;
  for(let i=0;i<buf.length;i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c>>>8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function zip(files){
  const local = [], central = [];
  let offset = 0;
  for(const f of files){
    const raw  = Buffer.from(f.data, "utf8");
    const comp = zlib.deflateRawSync(raw, {level:6});
    const name = Buffer.from(f.name, "utf8");
    const crc  = crc32(raw);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50,0); lh.writeUInt16LE(20,4); lh.writeUInt16LE(0x0800,6);
    lh.writeUInt16LE(8,8); lh.writeUInt16LE(0,10); lh.writeUInt16LE(0,12);
    lh.writeUInt32LE(crc,14); lh.writeUInt32LE(comp.length,18); lh.writeUInt32LE(raw.length,22);
    lh.writeUInt16LE(name.length,26); lh.writeUInt16LE(0,28);
    local.push(lh, name, comp);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50,0); ch.writeUInt16LE(20,4); ch.writeUInt16LE(20,6);
    ch.writeUInt16LE(0x0800,8); ch.writeUInt16LE(8,10); ch.writeUInt16LE(0,12); ch.writeUInt16LE(0,14);
    ch.writeUInt32LE(crc,16); ch.writeUInt32LE(comp.length,20); ch.writeUInt32LE(raw.length,24);
    ch.writeUInt16LE(name.length,28); ch.writeUInt16LE(0,30); ch.writeUInt16LE(0,32);
    ch.writeUInt16LE(0,34); ch.writeUInt16LE(0,36); ch.writeUInt32LE(0,38);
    ch.writeUInt32LE(offset,42);
    central.push(ch, name);
    offset += lh.length + name.length + comp.length;
  }
  const cd = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50,0); eocd.writeUInt16LE(0,4); eocd.writeUInt16LE(0,6);
  eocd.writeUInt16LE(files.length,8); eocd.writeUInt16LE(files.length,10);
  eocd.writeUInt32LE(cd.length,12); eocd.writeUInt32LE(offset,16); eocd.writeUInt16LE(0,20);
  return Buffer.concat([...local, cd, eocd]);
}
const xmlEsc = s => String(s==null?"":s)
  .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
  .replace(/"/g,"&quot;").replace(/\r?\n/g,"&#10;")
  .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,"");
function colName(n){
  let s = "";
  while(n > 0){ const m = (n-1) % 26; s = String.fromCharCode(65+m) + s; n = Math.floor((n-1)/26); }
  return s;
}
const STYLE = {NORMAL:0, HEADER:1, DATE:2, TIME:3};
function cellXml(ref, v, style){
  const s = style ? ` s="${style}"` : "";
  if(v == null || v === "") return "";
  if(typeof v === "number" && isFinite(v)) return `<c r="${ref}"${s}><v>${v}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
}
function sheetXml(header, rows){
  const cols = header.length;
  let out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
    + `<sheetViews><sheetView workbookViewId="0">`
    + `<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`
    + `<sheetFormatPr defaultRowHeight="15"/><cols>`;
  header.forEach((h,i)=>{
    const w = Math.min(46, Math.max(9, String(h).split("\n")[0].length + 3));
    out += `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;
  });
  out += `</cols><sheetData>`;
  out += `<row r="1" ht="34" customHeight="1">`
       + header.map((h,i)=> cellXml(colName(i+1)+"1", String(h), STYLE.HEADER)).join("")
       + `</row>`;
  rows.forEach((row, ri)=>{
    const r = ri + 2;
    out += `<row r="${r}">` + row.map((cell,i)=>{
      if(cell == null || cell === "") return "";
      const ref = colName(i+1) + r;
      if(cell && typeof cell === "object" && cell.t === "d") return cellXml(ref, cell.v, STYLE.DATE);
      if(cell && typeof cell === "object" && cell.t === "h") return cellXml(ref, cell.v, STYLE.TIME);
      return cellXml(ref, cell, STYLE.NORMAL);
    }).join("") + `</row>`;
  });
  out += `</sheetData><autoFilter ref="A1:${colName(cols)}${rows.length+1}"/></worksheet>`;
  return out;
}
const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="hh:mm"/></numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1D5FA8"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

function buildWorkbook(){
  const sheets = [];
  FORM_ORDER.forEach(fid=>{
    const name = SHEET[fid];
    sheets.push({name, header: HEADERS[name], rows: rowsFor(fid)});
  });
  sheets.push({name:"พนักงาน", header: HEADERS["พนักงาน"],
    rows: (M().employees).map((e,i)=> [i+1, e.name, e.nick, e.pos, e.tel])});
  sheets.push({name:"ผู้รับเหมา", header: HEADERS["ผู้รับเหมา"],
    rows: (M().contractors).map((c,i)=>["", i+1, c.code, c.name, c.addr, c.contact, c.plate, c.tel1, c.tel2, "", c.type, "", ""])});
  sheets.push({name:"ประเภทการจัดส่ง", header: HEADERS["ประเภทการจัดส่ง"], rows:[["BOX"],["SHEET"]]});

  const files = [];
  files.push({name:"[Content_Types].xml", data:
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
    + `<Default Extension="xml" ContentType="application/xml"/>`
    + `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`
    + sheets.map((s,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")
    + `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`
    + `</Types>`});
  files.push({name:"_rels/.rels", data:
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`});
  files.push({name:"xl/workbook.xml", data:
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`
    + sheets.map((s,i)=>`<sheet name="${xmlEsc(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")
    + `</sheets></workbook>`});
  files.push({name:"xl/_rels/workbook.xml.rels", data:
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + sheets.map((s,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("")
    + `<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`});
  files.push({name:"xl/styles.xml", data: STYLES_XML});
  sheets.forEach((s,i)=> files.push({name:`xl/worksheets/sheet${i+1}.xml`, data: sheetXml(s.header, s.rows)}));
  return zip(files);
}

/* ---------- แปลง record เป็นแถวตามคอลัมน์ของไฟล์เดิม ---------- */
function toSerial(iso){
  if(!iso) return "";
  const p = String(iso).split("-");
  if(p.length !== 3) return "";
  const ms = Date.UTC(+p[0], +p[1]-1, +p[2]);
  if(isNaN(ms)) return "";
  return Math.round((ms - Date.UTC(1899,11,30)) / 864e5);
}
function toFrac(hm){
  const m = String(hm||"").match(/^(\d{1,2}):(\d{2})/);
  if(!m) return "";
  return (Number(m[1])*60 + Number(m[2])) / 1440;
}
function rowsFor(fid){
  const map = COLMAP[SHEET[fid]] || [];
  const list = (DB.records[fid]||[]).slice().sort((a,b)=> String(a.date||"").localeCompare(String(b.date||"")) || (a.createdAt||0)-(b.createdAt||0));
  return list.map((r,idx)=>{
    const sc = calc(fid, r);
    return map.map(rule=>{
      if(!rule) return "";
      if(rule === "id")    return r.id || "";
      if(rule === "no")    return idx + 1;
      if(rule === "year")  return r.date ? Number(String(r.date).slice(0,4)) : "";
      if(rule === "month") return r.date ? Number(String(r.date).slice(5,7)) : "";
      if(rule === "date")  { const s = toSerial(r.date); return s===""? "" : {t:"d", v:s}; }
      if(rule === "time")  { const f = toFrac(r.time);   return f===""? "" : {t:"h", v:f}; }
      if(rule.startsWith("f:")){
        const k = rule.slice(2);
        const v = r[k];
        if(k === "verifyDate" || k === "approveDate"){ const s = toSerial(v); return s===""? "" : {t:"d", v:s}; }
        if((k==="verifier"||k==="inspector"||k==="approver") && !v) return r[k+"Name"] || "";
        return v == null ? "" : v;
      }
      if(rule.startsWith("i:")) { const v = (r.items||{})[rule.slice(2)]; return v == null ? "" : v; }
      if(rule.startsWith("c:")) { const v = sc[rule.slice(2)]; return v == null ? "" : v; }
      return "";
    });
  });
}

/* ============================================================
   HTTP
   ============================================================ */
const MIME = {".html":"text/html; charset=utf-8", ".js":"text/javascript; charset=utf-8",
  ".css":"text/css; charset=utf-8", ".json":"application/json; charset=utf-8",
  ".png":"image/png", ".jpg":"image/jpeg", ".jpeg":"image/jpeg", ".gif":"image/gif",
  ".svg":"image/svg+xml", ".xlsx":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".webmanifest":"application/manifest+json; charset=utf-8", ".ico":"image/x-icon",
  ".webp":"image/webp", ".txt":"text/plain; charset=utf-8"};

/* อนุญาตให้หน้าเว็บบน GitHub Pages เรียกเข้ามาได้ — ความปลอดภัยอยู่ที่โทเคน ไม่ใช่ที่ origin */
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
  "Access-Control-Max-Age": "86400"
};
function send(res, code, obj){
  const body = JSON.stringify(obj);
  res.writeHead(code, Object.assign({"Content-Type":"application/json; charset=utf-8", "Cache-Control":"no-store"}, CORS));
  res.end(body);
}
function readBody(req){
  return new Promise((resolve,reject)=>{
    let d = "";
    req.on("data", c=>{ d += c; if(d.length > 30e6) reject(new Error("ข้อมูลใหญ่เกินไป")); });
    req.on("end", ()=>{ try{ resolve(d? JSON.parse(d) : {}); }catch(e){ reject(e); } });
    req.on("error", reject);
  });
}

/* ที่อยู่ที่เครื่องอื่นในวง LAN ใช้เปิดแอปได้ */
function lanIPs(){
  const out = [];
  Object.values(os.networkInterfaces()).forEach(list=> (list||[]).forEach(nic=>{
    if(nic.family === "IPv4" && !nic.internal) out.push(nic.address);
  }));
  return out;
}
function lanURLs(){
  const ips = lanIPs();
  /* ถ้าเปิด HTTPS ไว้ ให้เสนอที่อยู่ https ก่อน — เป็นทางเดียวที่ Android ติดตั้งเป็นแอปได้เต็มรูปแบบ */
  const https = HTTPS_SERVER ? ips.map(ip=> "https://" + ip + ":" + HTTPS_PORT + "/") : [];
  return https.concat(ips.map(ip=> "http://" + ip + ":" + PORT + "/"));
}

const handler = async (req,res)=>{
  const url = decodeURIComponent(req.url.split("?")[0]);
  try{
    /* ---- preflight ของเบราว์เซอร์ ---- */
    if(req.method === "OPTIONS"){ res.writeHead(204, CORS); return res.end(); }

    /* ---- ต้องใส่รหัสไหม: ในออฟฟิศ (localhost/LAN) ไม่ต้อง, จากอุโมงค์ต้องใส่ ---- */
    const viaTunnel = !!(req.headers["cf-connecting-ip"] || req.headers["cf-ray"]);
    const needsAuth = viaTunnel || (CFG.requireCodeOnLan && !isLocalReq(req));

    if(url === "/api/ping") return send(res, 200, {ok:true, mode:"server", file:path.basename(DATA_XLSX),
                                                   urls:lanURLs(), https:!!HTTPS_SERVER,
                                                   needsAuth, publicUrl:PUBLIC_URL || "",
                                                   hasBase:!!BASE});

    if(url === "/api/login" && req.method === "POST"){
      const ip = clientIP(req);
      const wait = loginBlocked(ip);
      if(wait) return send(res, 429, {ok:false, error:"ใส่รหัสผิดหลายครั้ง กรุณารออีก " + Math.ceil(wait/60) + " นาที"});
      const b = await readBody(req);
      const given = String((b && b.code) || "").trim();
      const want  = String(CFG.accessCode);
      const ok = given.length === want.length &&
                 crypto.timingSafeEqual(Buffer.from(given.padEnd(32)), Buffer.from(want.padEnd(32)));
      if(!ok){ noteLoginFail(ip); return send(res, 401, {ok:false, error:"รหัสไม่ถูกต้อง"}); }
      LOGIN_FAILS.delete(ip);
      console.log("  เข้าสู่ระบบสำเร็จจาก " + ip);
      return send(res, 200, {ok:true, token:makeToken(30)});
    }

    /* ---- ทุกเส้นทางที่เหลือของ /api/ ต้องมีโทเคน ถ้ามาจากนอกออฟฟิศ ---- */
    if(url.startsWith("/api/") && needsAuth){
      const auth = String(req.headers.authorization || "");
      const tok  = auth.startsWith("Bearer ") ? auth.slice(7) : "";
      if(!validToken(tok)) return send(res, 401, {ok:false, error:"ต้องใส่รหัสเข้าใช้งานก่อน", needsAuth:true});
    }

    if(url === "/api/bootstrap"){
      if(!BASE) return send(res, 503, {ok:false, error:"ยังไม่มีไฟล์ WT-Paperless-Base.json"});
      const buf = bootstrapGz();
      res.writeHead(200, Object.assign({"Content-Type":"application/json; charset=utf-8",
                                        "Content-Encoding":"gzip", "Cache-Control":"no-store"}, CORS));
      return res.end(buf);
    }
    if(url === "/api/state") return send(res, 200, {ok:true, records:DB.records, master:DB.master, savedAt:DB.savedAt});

    if(url === "/api/master" && req.method === "POST"){
      const b = await readBody(req);
      const m = b && b.master;
      if(!m || (!Array.isArray(m.contractors) && !Array.isArray(m.employees)))
        return send(res, 400, {ok:false, error:"ข้อมูลหลักไม่ถูกต้อง"});
      DB.master = {
        contractors: Array.isArray(m.contractors) ? m.contractors : M().contractors,
        employees:   Array.isArray(m.employees)   ? normEmpList(m.employees) : M().employees,
        plates:      Array.isArray(m.plates)      ? m.plates      : undefined
      };
      persist();
      console.log(`  อัปเดตข้อมูลหลัก — ผู้รับเหมา ${DB.master.contractors.length} ราย · พนักงาน ${DB.master.employees.length} คน`);
      return send(res, 200, {ok:true, master:DB.master});
    }

    if(url === "/api/save" && req.method === "POST"){
      const b = await readBody(req);
      const fid = b.fid, list = Array.isArray(b.records) ? b.records : [b.record];
      if(!SHEET[fid]) return send(res, 400, {ok:false, error:"ไม่รู้จักแบบฟอร์มนี้"});
      const saved = [];
      list.forEach(rec=>{
        if(!rec) return;
        if(!rec.id) rec.id = Date.now().toString(36) + Math.random().toString(36).slice(2,7);
        if(!rec.createdAt) rec.createdAt = Date.now();
        rec.updatedAt = Date.now();
        saveSignatures(rec);
        const arr = DB.records[fid];
        const i = arr.findIndex(x=> x.id === rec.id);
        if(i >= 0) arr[i] = rec; else arr.push(rec);
        saved.push(rec);
      });
      persist();
      console.log(`  บันทึก ${saved.length} รายการ (${SHEET[fid]}) — รวมทั้งหมด ${DB.records[fid].length} รายการ`);
      return send(res, 200, {ok:true, records:saved});
    }

    if(url === "/api/delete" && req.method === "POST"){
      const b = await readBody(req);
      if(!SHEET[b.fid]) return send(res, 400, {ok:false, error:"ไม่รู้จักแบบฟอร์มนี้"});
      const before = DB.records[b.fid].length;
      DB.records[b.fid] = DB.records[b.fid].filter(x=> x.id !== b.id);
      persist();
      return send(res, 200, {ok:true, removed: before - DB.records[b.fid].length});
    }

    /* ไฟล์ปกติ */
    let rel = url === "/" ? APP_FILE : url.replace(/^\//,"");
    const file = path.join(ROOT, rel);
    if(!file.startsWith(ROOT)) { res.writeHead(403); return res.end("forbidden"); }
    if(!fs.existsSync(file) || fs.statSync(file).isDirectory()){ res.writeHead(404); return res.end("ไม่พบไฟล์"); }
    if(path.basename(file) === "WT-Paperless-Base.json" || path.basename(file) === "WT-Paperless-Config.json"){
      res.writeHead(403); return res.end("forbidden");   /* ข้อมูลดิบต้องผ่าน /api/bootstrap เท่านั้น */
    }
    res.writeHead(200, Object.assign({"Content-Type": MIME[path.extname(file).toLowerCase()] || "application/octet-stream",
                        "Cache-Control": path.extname(file)===".html" ? "no-store" : "max-age=3600"}, CORS));
    fs.createReadStream(file).pipe(res);
  }catch(e){
    console.log("  ! ผิดพลาด:", e.message);
    send(res, 500, {ok:false, error:e.message});
  }
};

const server = http.createServer(handler);

/* ============================================================
   อุโมงค์ Cloudflare — ให้เข้าถึงจากนอกออฟฟิศโดยไม่ต้องเปิดพอร์ตที่เราเตอร์
   ต้องมี cloudflared.exe วางไว้ข้างไฟล์นี้ (หรือติดตั้งไว้ใน PATH)
   ============================================================ */
let PUBLIC_URL = "", TUNNEL = null;
function cloudflaredPath(){
  const local = path.join(ROOT, "cloudflared.exe");
  if(fs.existsSync(local)) return local;
  return process.env.WT_CLOUDFLARED || "cloudflared";
}
function startTunnel(onUrl){
  if(process.env.WT_NO_TUNNEL) return null;
  let proc;
  try{
    proc = spawn(cloudflaredPath(), ["tunnel","--no-autoupdate","--url","http://localhost:" + PORT],
                 {windowsHide:true});
  }catch(e){ return null; }
  proc.on("error", ()=>{
    console.log("  ! ไม่พบ cloudflared — ใช้ได้เฉพาะในวง Wi-Fi เดียวกัน");
    console.log("    (ดาวน์โหลด cloudflared.exe มาวางข้างไฟล์นี้ แล้วเปิดใหม่)");
  });
  const scan = buf=>{
    const m = String(buf).match(/https:\/\/[a-z0-9-]+\.trycloudflare\.com/i);
    if(m && !PUBLIC_URL){
      PUBLIC_URL = m[0];
      if(onUrl) onUrl(PUBLIC_URL);
    }
  };
  proc.stdout.on("data", scan);
  proc.stderr.on("data", scan);
  return proc;
}
function shareLink(){
  if(!PUBLIC_URL) return "";
  const page = (CFG && CFG.pagesUrl) || "";
  return page ? (page.replace(/\/*$/, "/") + "?api=" + encodeURIComponent(PUBLIC_URL)) : PUBLIC_URL;
}

/* ---- HTTPS (ไม่บังคับ) ----
   วาง cert.pem + key.pem ไว้ในโฟลเดอร์ WT-Paperless-Cert แล้วระบบจะเปิดพอร์ต 8443 ให้เอง
   จำเป็นเฉพาะกรณีอยากให้ Android ติดตั้งเป็นแอปเต็มรูปแบบ (ต้องติดตั้งใบรับรองในมือถือด้วย) */
const CERT_DIR   = path.join(ROOT, "WT-Paperless-Cert");
const HTTPS_PORT = Number(process.env.WT_HTTPS_PORT || 8443);
let HTTPS_SERVER = null;
function startHTTPS(){
  const key = path.join(CERT_DIR, "key.pem"), crt = path.join(CERT_DIR, "cert.pem");
  if(!fs.existsSync(key) || !fs.existsSync(crt)) return null;
  try{
    const https = require("https");
    const srv = https.createServer({key:fs.readFileSync(key), cert:fs.readFileSync(crt)}, handler);
    srv.listen(HTTPS_PORT, "0.0.0.0");
    srv.on("error", e=> console.log("  ! เปิด HTTPS ไม่ได้:", e.message));
    return srv;
  }catch(e){ console.log("  ! เปิด HTTPS ไม่ได้:", e.message); return null; }
}

/* ============================================================
   เริ่มทำงาน
   ============================================================ */
loadConfig();
loadBase();
loadDB();
HTTPS_SERVER = startHTTPS();
const ips = lanIPs();
server.listen(PORT, "0.0.0.0", ()=>{
  const total = FORM_ORDER.reduce((a,f)=> a + DB.records[f].length, 0);
  console.log("");
  console.log("  ============================================================");
  console.log("   WT Paperless — ระบบตรวจสอบ คลังสินค้าและจัดส่ง");
  console.log("  ============================================================");
  console.log("");
  console.log("   เปิดใช้งานที่เครื่องนี้ :  http://localhost:" + PORT);
  ips.forEach(ip=> console.log("   เปิดจากมือถือ/เครื่องอื่น :  http://" + ip + ":" + PORT));
  if(HTTPS_SERVER) ips.forEach(ip=> console.log("   แบบ HTTPS (ติดตั้งเป็นแอปได้เต็มที่) :  https://" + ip + ":" + HTTPS_PORT));
  console.log("");
  console.log("   รหัสเข้าใช้งานจากนอกออฟฟิศ :  " + CFG.accessCode + "   (แก้ได้ในไฟล์ WT-Paperless-Config.json)");
  if(!BASE) console.log("   ! ไม่พบ WT-Paperless-Base.json — แอปจะไม่มีข้อมูลย้อนหลังให้ดู");

  TUNNEL = startTunnel(u=>{
    console.log("");
    console.log("  ============================================================");
    console.log("   เปิดจากที่ไหนก็ได้ (ไม่ต้องอยู่ Wi-Fi เดียวกัน)");
    console.log("  ============================================================");
    const link = shareLink();
    console.log("   ที่อยู่ :  " + link);
    console.log("   รหัส   :  " + CFG.accessCode);
    console.log("");
    console.log("   * ที่อยู่นี้เปลี่ยนใหม่ทุกครั้งที่เปิดโปรแกรม");
    console.log("     ให้เปิดเมนู \"ติดตั้งเป็นแอป\" ในแอป แล้วให้พนักงานสแกน QR ใหม่");
  });
  if(TUNNEL) console.log("   กำลังเปิดอุโมงค์สำหรับใช้งานนอกออฟฟิศ…");
  console.log("");
  console.log("   ข้อมูลใหม่บันทึกลง :  " + path.basename(DATA_XLSX));
  console.log("   สำรองข้อมูลอัตโนมัติที่ :  " + path.basename(BACKUP) + "\\");
  console.log("   มีข้อมูลที่บันทึกไว้แล้ว :  " + total + " รายการ");
  console.log("");
  console.log("   * ไฟล์ Paperless of Warehouse & transportation.xlsx เดิม ไม่ถูกแตะต้อง");
  console.log("   * ปิดระบบ: กดปิดหน้าต่างนี้ หรือกด Ctrl+C");
  console.log("");
  if(!process.env.WT_NOOPEN){
    try{ require("child_process").exec('start "" "http://localhost:' + PORT + '"'); }catch(e){}
  }
});
server.on("error", e=>{
  if(e.code === "EADDRINUSE"){
    console.log("\n  ! พอร์ต " + PORT + " ถูกใช้งานอยู่แล้ว — ระบบอาจเปิดค้างอยู่ ลองเปิด http://localhost:" + PORT + " ดูก่อน\n");
  } else console.log("\n  ! เปิดระบบไม่สำเร็จ:", e.message, "\n");
});
