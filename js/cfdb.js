/*
 * แทนที่ js/firestore.js — คุยกับ Cloudflare Worker + D1 (cf-worker/) แทน Firebase Firestore
 * เก็บชื่อฟังก์ชัน (window.fsList/fsGet/fsCreate/fsUpdate/fsUpsert/fsDelete) เหมือนเดิมทุกตัว
 * เพื่อให้ index.html/desktop.html ไม่ต้องแก้โค้ดส่วนเรียกใช้เลย — แก้แค่ตรง <script> ที่โหลดไฟล์นี้แทน
 * ต้องโหลดไฟล์นี้ "หลัง" cf-config.js เสมอ (ใช้ CF_API_BASE จากไฟล์นั้น)
 *
 * ไม่มี auth/token เหมือนเดิมที่เคยมี anonymous auth — endpoint นี้เปิดกว้างตามที่ตกลงกันไว้
 * (เทียบเท่าของเดิมที่ Firestore Rules อนุญาตให้ทุกคนที่ล็อกอินแบบ anonymous อ่าน/เขียนได้อยู่แล้ว)
 */

function cfUrl(path) {
  return CF_API_BASE.replace(/\/$/, '') + '/' + path;
}

// สัญญาณเน็ตในร้านมักอ่อน/หลุดกลางทาง — fetch() เปล่าๆ ในสภาพนั้นจะ "ค้าง" ไม่ resolve ไม่ reject
// เป็นนาที ทำให้แอปติดอยู่ที่ "กำลังตรวจสอบ…"/หน้าโหลดข้อมูลตลอดไป ไม่ยอม fallback ไปโหมดออฟไลน์
// ตามที่ออกแบบไว้ (เทียบกับ error แบบขาดการเชื่อมต่อจริงๆ ที่ fetch reject เร็วอยู่แล้ว) ใส่ timeout
// เองเพื่อให้ "เน็ตค้าง" พฤติกรรมเหมือน "ไม่มีเน็ต" คือ fail เร็วแล้วปล่อยให้โค้ดเดิม fallback ตามปกติ
var CF_FETCH_TIMEOUT_MS = 8000;
function cfFetch(url, opts) {
  var ctrl = new AbortController();
  var timer = setTimeout(function () { ctrl.abort(); }, CF_FETCH_TIMEOUT_MS);
  var merged = Object.assign({}, opts, { signal: ctrl.signal });
  return fetch(url, merged).finally(function () { clearTimeout(timer); });
}

// เดิม index.html/desktop.html เรียก fsAutoLogin() เพื่อล็อกอิน Firebase Anonymous Auth ก่อนใช้งาน
// ที่นี่ไม่มี auth ให้ล็อกอิน แต่คงชื่อฟังก์ชันไว้ให้เรียกได้เหมือนเดิม — ใช้เป็น health check
// เช็คว่าเข้าถึง Worker ได้จริง (ถ้าเข้าไม่ได้ throw เหมือนเดิม ให้ตกไป catch(e){S.online=false})
window.fsAutoLogin = async function () {
  const r = await cfFetch(cfUrl(''));
  if (!r.ok) throw new Error('Worker unreachable: ' + r.status);
  return true;
};

window.fsList = async function (col) {
  const r = await cfFetch(cfUrl(col));
  const d = await r.json();
  if (d.error) throw new Error(d.error.message);
  return d.documents || [];
};

window.fsGet = async function (col, id) {
  const r = await cfFetch(cfUrl(col + '/' + encodeURIComponent(id)));
  if (r.status === 404) return null;
  const d = await r.json();
  if (d.error) throw new Error(d.error.message);
  return d;
};

window.fsCreate = async function (col, id, obj) {
  const r = await cfFetch(cfUrl(col + '/' + encodeURIComponent(id)), {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(obj)
  });
  const d = await r.json();
  if (d.error) {
    const err = new Error(d.error.message);
    err.alreadyExists = r.status === 409 || /already exists/i.test(d.error.message || '');
    throw err;
  }
  return d;
};

// อัปเดตเฉพาะฟิลด์ที่ระบุ (merge ฝั่ง Worker ให้แล้ว)
window.fsUpdate = async function (col, id, obj) {
  const r = await cfFetch(cfUrl(col + '/' + encodeURIComponent(id)), {
    method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(obj)
  });
  const d = await r.json();
  if (d.error) throw new Error(d.error.message);
  return d;
};

// สร้างหรือแทนที่ทั้ง document (ใช้ POST ก่อน ถ้ามีอยู่แล้วให้ fallback เป็น fsUpdate เต็มฟิลด์)
window.fsUpsert = async function (col, id, obj) {
  try {
    return await window.fsCreate(col, id, obj);
  } catch (e) {
    return await window.fsUpdate(col, id, obj);
  }
};

window.fsDelete = async function (col, id) {
  const r = await cfFetch(cfUrl(col + '/' + encodeURIComponent(id)), { method: 'DELETE' });
  const d = await r.json();
  if (d.error) throw new Error(d.error.message);
  return d;
};

// ไม่มีการเรียกใช้จริงในแอปตอนนี้ (เช็คแล้วทั้ง index.html/desktop.html ไม่ได้ใช้ fsQuery) — ใส่ไว้เผื่ออนาคต
window.fsQuery = async function () {
  throw new Error('fsQuery: not supported by Cloudflare backend, use fsList + filter client-side');
};
