/*
 * Cloudflare Worker — แทนที่ Firestore REST API เดิม (js/firestore.js)
 * เก็บข้อมูลใน D1 เป็นตาราง docstore เดียว (col, id, data JSON, updated_at)
 * เลียนแบบ contract เดิมของ window.fsList/fsGet/fsCreate/fsUpdate/fsUpsert ให้ตรงที่สุด
 * เพื่อให้ js/cfdb.js (ฝั่ง client) แทนที่ js/firestore.js ได้โดยโค้ดส่วนอื่นแทบไม่ต้องแก้
 *
 * ไม่มีการเช็คสิทธิ์ผู้เรียก (ตามที่ตกลง — ไม่มีรหัสผ่านเหมือนระบบเดิมที่ใช้ Firebase Anonymous Auth
 * ผ่าน Rules แบบเปิดกว้างอยู่แล้ว) endpoint นี้เปิดให้เรียกได้จากทุกที่ที่รู้ URL
 *
 * ผูก D1 binding ชื่อ DB ใน wrangler.toml ก่อน deploy
 */

function cors(resp) {
  resp.headers.set('Access-Control-Allow-Origin', '*');
  resp.headers.set('Access-Control-Allow-Methods', 'GET,POST,PATCH,DELETE,OPTIONS');
  resp.headers.set('Access-Control-Allow-Headers', 'Content-Type');
  return resp;
}
function json(obj, status) {
  return cors(new Response(JSON.stringify(obj), {
    status: status || 200,
    headers: { 'Content-Type': 'application/json' }
  }));
}
function rowToDoc(row) {
  const o = JSON.parse(row.data);
  o._id = row.id;
  return o;
}

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') return cors(new Response(null, { status: 204 }));

    const url = new URL(request.url);
    const parts = url.pathname.split('/').filter(Boolean); // [col] หรือ [col, id] หรือ ['admin','import']

    try {
      if (parts[0] === 'admin' && parts[1] === 'import' && request.method === 'POST') {
        // ใช้ครั้งเดียวตอนย้ายข้อมูลจาก Firestore เดิม — { col, rows: [{id, data}] }
        const body = await request.json();
        const stmt = env.DB.prepare(
          'INSERT INTO docstore (col, id, data, updated_at) VALUES (?, ?, ?, ?) ' +
          'ON CONFLICT(col, id) DO UPDATE SET data=excluded.data, updated_at=excluded.updated_at'
        );
        const now = Date.now();
        const batch = (body.rows || []).map(r => stmt.bind(body.col, String(r.id), JSON.stringify(r.data), now));
        if (batch.length) await env.DB.batch(batch);
        return json({ imported: batch.length });
      }

      const col = parts[0];
      const id = parts[1];
      if (!col) return json({ ok: true }); // health check — ใช้แทน fsAutoLogin เดิมของ Firebase (เช็คว่าเข้าถึง API ได้)

      if (request.method === 'GET' && !id) {
        // fsList — ปกติคืนทุก document ใน collection แต่ documents รองรับ ?unprinted=1 ให้กรองเฉพาะ
        // เอกสารที่ status ยังไม่เป็น "printed" ตั้งแต่ชั้น SQL เลย (ลดจำนวน row read ต่อครั้ง ไม่ต้อง
        // ดึงทั้งหมดมาแล้วมากรองฝั่ง client) — เดสก์ท็อปใช้ตอน poll เป็นระยะ เอกสารที่ printed แล้วโหลด
        // เต็มครั้งแรกครั้งเดียวพอ ไม่ต้องดึงซ้ำทุกรอบ
        let sql = 'SELECT id, data FROM docstore WHERE col = ?';
        if (col === 'documents' && url.searchParams.get('unprinted') === '1') {
          sql += " AND json_extract(data, '$.status') != 'printed'";
        }
        const { results } = await env.DB.prepare(sql).bind(col).all();
        return json({ documents: results.map(rowToDoc) });
      }

      if (request.method === 'GET' && id) {
        // fsGet
        const row = await env.DB.prepare('SELECT id, data FROM docstore WHERE col = ? AND id = ?').bind(col, id).first();
        if (!row) return json({ error: { message: 'NOT_FOUND' } }, 404);
        return json(rowToDoc(row));
      }

      if (request.method === 'POST') {
        // fsCreate — ต้องไม่มี id นี้อยู่ก่อน (409 ถ้ามีแล้ว เหมือน Firestore documentId ซ้ำ)
        const docId = id || url.searchParams.get('documentId');
        if (!docId) return json({ error: { message: 'missing documentId' } }, 400);
        const existing = await env.DB.prepare('SELECT id FROM docstore WHERE col = ? AND id = ?').bind(col, docId).first();
        if (existing) return json({ error: { message: 'ALREADY_EXISTS: document already exists' } }, 409);
        const data = await request.json();
        await env.DB.prepare('INSERT INTO docstore (col, id, data, updated_at) VALUES (?, ?, ?, ?)')
          .bind(col, docId, JSON.stringify(data), Date.now()).run();
        const o = Object.assign({}, data); o._id = docId;
        return json(o);
      }

      if (request.method === 'PATCH' && id) {
        // fsUpdate — merge เฉพาะฟิลด์ที่ส่งมาเข้ากับของเดิม (เหมือน updateMask ของ Firestore)
        const row = await env.DB.prepare('SELECT data FROM docstore WHERE col = ? AND id = ?').bind(col, id).first();
        if (!row) return json({ error: { message: 'NOT_FOUND' } }, 404);
        const patch = await request.json();
        const merged = Object.assign(JSON.parse(row.data), patch);
        await env.DB.prepare('UPDATE docstore SET data = ?, updated_at = ? WHERE col = ? AND id = ?')
          .bind(JSON.stringify(merged), Date.now(), col, id).run();
        const o = Object.assign({}, merged); o._id = id;
        return json(o);
      }

      if (request.method === 'DELETE' && id) {
        // อนุญาตลบจริงเฉพาะ collection "documents" ตามที่ขอ (ผู้ดูแลลบเอกสารทิ้งถาวรได้จากหน้าเดสก์ท็อป)
        // collection อื่น (employees ฯลฯ) ยังบล็อกไว้เหมือนเดิม — ใช้ soft-delete (fsUpdate ตั้ง active:false)
        if (col !== 'documents') {
          return json({ error: { message: 'PERMISSION_DENIED: delete disabled for this collection, use soft-delete' } }, 403);
        }
        const existing = await env.DB.prepare('SELECT id FROM docstore WHERE col = ? AND id = ?').bind(col, id).first();
        if (!existing) return json({ error: { message: 'NOT_FOUND' } }, 404);
        await env.DB.prepare('DELETE FROM docstore WHERE col = ? AND id = ?').bind(col, id).run();
        return json({ ok: true });
      }

      return json({ error: { message: 'unsupported route' } }, 404);
    } catch (e) {
      return json({ error: { message: String(e && e.message || e) } }, 500);
    }
  }
};
