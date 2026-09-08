/*
 * สร้างไฟล์ Excel จากเทมเพลตต้นฉบับ (templates/ADJ_template.xlsx, templates/RTC_template.xlsx)
 * ใช้ร่วมกันทั้งฝั่งมือถือ (index.html) และเดสก์ท็อป (desktop.html)
 * ต้องโหลด vendor/exceljs.min.js (ตัวแปร ExcelJS) และ vendor/jszip.min.js (ตัวแปร JSZip)
 * ให้พร้อมก่อนเรียกฟังก์ชันในไฟล์นี้ — ExcelJS ใช้แก้ workbook เข้าใจโครงสร้างจริง (แทนที่การ
 * แก้ xl/worksheets/sheet2.xml ด้วย regex แบบเดิม ซึ่งเปราะกับ merged cells/styles/formula
 * — พิสูจน์แล้วด้วย POC ว่า ExcelJS รักษา style/merge/print setting/สูตรได้ครบ ไม่ต้องคำนวณ
 * Excel date serial หรือแมพ style คู่แฝดเองแบบตอนใช้ regex)
 * JSZip ใช้แค่ห่อไฟล์ .xlsx หลายหน้าเป็น .zip เดียวตอนเอกสารเกิน cap ต่อหน้า (buildExportPackage)
 *
 * สำคัญ: RTC_template.xlsx คอลัมน์ H (ราคาใหม่) เป็นสูตร Excel จริง
 * (IFERROR(ROUNDUP(E-(E*F),0),"")) — ห้ามเขียนค่าทับคอลัมน์นี้ ปล่อยให้ Excel คำนวณเองตอนเปิดไฟล์
 */

var EXPORT_CFG = {
  ADJ: { sheetName: 'adjust_qty_form1', first: 5, last: 47, hdrBranch: 'A2', hdrDate: 'A3',
         cols: { no: 'A', bc: 'B', name: 'C', code: 'D', qty: 'E', reason: 'F', exp: 'G' },
         templateUrl: 'templates/ADJ_template.xlsx?v=20260905-2000' },
  RTC: { sheetName: 'From', first: 9, last: 28, hdrBranch: 'A4', hdrDate: 'A5',
         cols: { no: 'A', bc: 'B', name: 'C', unit: 'D', price: 'E', pct: 'F', qty: 'G',
                 newPrice: 'H', left: 'I', dateC25: 'J', exp: 'K' },
         templateUrl: 'templates/RTC_template.xlsx?v=20260905-2000' },
  // ฟอร์ม "Daily Running Number Report" (ติดตามบันทึกรับสินค้าตาม PO/Supplier) — ไม่มีสูตร,
  // ไม่มีบาร์โค้ด, กรอกมือล้วน คนละชุดคอลัมน์กับ ADJ/RTC โดยสิ้นเชิง
  TRACK: { sheetName: 'TrackingForm', first: 10, last: 34, hdrBranch: 'A6', hdrDate: 'A5',
           cols: { no: 'A', runNo: 'B', po: 'C', supNo: 'D', supName: 'E', recvDate: 'F',
                   confirmDone: 'G', confirmPending: 'H', remark: 'I' },
           templateUrl: 'templates/TRACK_template.xlsx?v=20260907-1031' }
};

function exportThDate(iso) { if (!iso) return ''; const p = iso.split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : iso; }
// Running No. รูปแบบ YYMMDDNNNN ตามหัวตาราง ("Running No.") ของฟอร์ม Tracking — ปีย่อ 2 หลัก+เดือน+วัน
// ของเอกสาร ต่อด้วยลำดับรายการ 4 หลัก เริ่มที่ 0001 (รายการแรก) ให้ตรงกับตราปั๊มเลขจริงที่ร้านใช้
function exportTrackRunningNo(dateIso, seq1) {
  const p = String(dateIso || '').split('-');
  if (p.length !== 3) return '';
  return p[0].slice(-2) + p[1] + p[2] + String(seq1).padStart(4, '0');
}
// ตัดอักขระที่ใช้เป็นชื่อไฟล์ไม่ได้ (Windows/macOS ต้องห้าม \/:*?"<>| และช่องว่างต้นท้าย)
function exportSanitizeFilename(s) {
  return String(s || '').trim().replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, '');
}
// ห้ามเซ็ต cell.numFmt ตรงๆ — ExcelJS เก็บ style ของหลายเซลล์ที่หน้าตาเหมือนกันไว้เป็น object
// เดียวกันโดยอ้างอิงร่วมกัน (เช่น A47/D47/E47/G47 ในเทมเพลตนี้ใช้ style เดียวกันทั้งแถว) ถ้าแก้
// cell.numFmt ตรงๆ จะไปเปลี่ยน numFmt ของทุกเซลล์ที่ใช้ style เดียวกันนั้นด้วยโดยไม่ตั้งใจ
// (พบจริงตอนทดสอบ: ตั้งวันที่ให้ G47 แล้ว NO./Adjust code/จำนวน ในแถวเดียวกันกลายเป็นวันที่ไปด้วย)
// ต้อง clone style เป็น object ใหม่ก่อนแก้ ให้กระทบเฉพาะเซลล์นี้เซลล์เดียว
function exportSetNumFmt(cell, numFmt) {
  cell.style = Object.assign({}, cell.style, { numFmt: numFmt });
}
// เขียนวันหมดอายุเป็น Excel date จริง (ไม่ใช่ string) — ExcelJS ตั้ง number format ต่อเซลล์ได้ตรงๆ
// ไม่ต้องคำนวณ Excel serial เองหรือแมพ style คู่แฝดแบบตอนใช้ regex patch XML
function exportSetDateCell(cell, iso) {
  if (!iso) { cell.value = null; return; }
  const p = String(iso).split('-');
  if (p.length !== 3) { cell.value = null; return; }
  const y = Number(p[0]), mo = Number(p[1]), d = Number(p[2]);
  if (!y || !mo || !d) { cell.value = null; return; }
  cell.value = new Date(Date.UTC(y, mo - 1, d));
  exportSetNumFmt(cell, 'dd/mm/yyyy');
}

var exportTemplateCache = {};
async function exportLoadTemplate(mode) {
  if (exportTemplateCache[mode]) return exportTemplateCache[mode];
  // ห้ามใช้ไฟล์ template เก่าจาก browser/CDN เพราะเคยทำให้รูปแบบคอลัมน์ไม่ตรงต้นฉบับ
  const res = await fetch(EXPORT_CFG[mode].templateUrl, { cache: 'no-store' });
  if (!res.ok) throw new Error('โหลดแบบฟอร์มต้นฉบับไม่สำเร็จ (' + res.status + ')');
  const buf = await res.arrayBuffer();
  exportTemplateCache[mode] = buf;
  return buf;
}

// บั๊กของ ExcelJS: โลโก้ (xl/drawings/drawing1.xml) ที่โหลดมาจาก template แล้ว save กลับ
// จะถูกเขียน <a:off>/<a:ext> (ตำแหน่ง/ขนาดรูป) เป็น 0 ทั้งหมด — Excel เปิดแล้วขึ้น [Repaired]
// "Removed Part: /xl/drawings/drawing1.xml (Drawing shape)" เพราะรูปที่ขนาด 0 ผิดสเปก OOXML
// แก้โดยจำค่าตำแหน่ง/ขนาดจริงจาก template ต้นฉบับไว้ก่อน แล้วแปะกลับเข้าไปในไฟล์ที่ ExcelJS เขียนออกมา
var exportDrawingXfrmCache = {};
async function exportGetOriginalDrawingXfrm(mode, templateBytes) {
  if (mode in exportDrawingXfrmCache) return exportDrawingXfrmCache[mode];
  let result = null;
  try {
    const zip = await JSZip.loadAsync(templateBytes.slice(0));
    const f = zip.file('xl/drawings/drawing1.xml');
    if (f) {
      const xml = await f.async('string');
      const m = xml.match(/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
      if (m) result = { x: m[1], y: m[2], cx: m[3], cy: m[4] };
    }
  } catch (e) { /* ไม่มีรูปในเทมเพลตนี้ก็ข้ามไป ไม่ถือเป็น error */ }
  exportDrawingXfrmCache[mode] = result;
  return result;
}
async function exportFixDrawingXfrm(xlsxBuf, mode, templateBytes) {
  const xfrm = await exportGetOriginalDrawingXfrm(mode, templateBytes);
  if (!xfrm) return xlsxBuf;
  const zip = await JSZip.loadAsync(xlsxBuf);
  const f = zip.file('xl/drawings/drawing1.xml');
  if (!f) return xlsxBuf;
  let xml = await f.async('string');
  if (!/<a:off x="0" y="0"\/><a:ext cx="0" cy="0"\/>/.test(xml)) return xlsxBuf; // ExcelJS แก้บั๊กนี้แล้วในเวอร์ชันใหม่กว่า ไม่ต้องแตะ
  xml = xml.replace('<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>',
    '<a:off x="' + xfrm.x + '" y="' + xfrm.y + '"/><a:ext cx="' + xfrm.cx + '" cy="' + xfrm.cy + '"/>');
  zip.file('xl/drawings/drawing1.xml', xml);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

/*
 * mode: 'ADJ' | 'RTC'
 * items: array ของรายการ (โครงสร้างเดียวกับ S.items[mode] เดิม)
 * meta: {branch, dept, date, docId, employeeCode} (date เป็น 'YYYY-MM-DD')
 * page: เลขหน้า (0 = หน้าแรก) — 1 หน้ากรอกได้ cap แถวตาม cfg
 * คืนค่า {blob, filename}
 */
window.buildExportXlsx = async function (mode, items, meta, page) {
  const cfg = EXPORT_CFG[mode], cap = cfg.last - cfg.first + 1;
  const pageItems = items.slice(page * cap, page * cap + cap);

  // โหลด workbook ใหม่จาก bytes ต้นฉบับทุกครั้ง ไม่ใช้ object ร่วมกันระหว่างเอกสาร
  const templateBytes = await exportLoadTemplate(mode);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateBytes.slice(0));
  const ws = wb.getWorksheet(cfg.sheetName);
  if (!ws) throw new Error('แบบฟอร์มต้นฉบับไม่สมบูรณ์ (ไม่พบชีท ' + cfg.sheetName + ')');

  const branch = (meta.branch || '').trim(), dept = (meta.dept || '').trim();
  const d = meta.date ? exportThDate(meta.date) : '';
  if (mode === 'TRACK') {
    // ฟอร์มนี้ไม่มีช่อง "แผนก" ในหัวเอกสาร ต่างจาก ADJ/RTC
    ws.getCell(cfg.hdrBranch).value = 'สาขา ' + (branch || '.......................');
    ws.getCell(cfg.hdrDate).value = 'วันที่ ' + (d || '.......................');
  } else {
    ws.getCell(cfg.hdrBranch).value = 'สาขา ' + (branch || '.......................') +
      '     แผนก ' + (dept || '.......................');
    ws.getCell(cfg.hdrDate).value = 'วันที่ ' + (d || '.......................');
  }

  const C = cfg.cols;
  for (let i = 0; i < cap; i++) {
    const r = cfg.first + i, it = pageItems[i];
    if (!it) {
      for (const k in C) {
        if (mode === 'RTC' && k === 'newPrice') continue; // ปล่อยสูตรเดิมไว้ ไม่แตะ
        ws.getCell(C[k] + r).value = null;
      }
      continue;
    }
    ws.getCell(C.no + r).value = page * cap + i + 1;
    if (mode === 'TRACK') {
      // meta.trackStartSeq = จำนวนรายการของเอกสาร TRACK อื่นๆ ในวันเดียวกันที่ "มาก่อน" เอกสารนี้
      // (เดสก์ท็อปคำนวณให้ก่อนเรียก) ไม่งั้น Running No. จะรีเซ็ตเริ่ม 1 ใหม่ทุกครั้งที่เปิดเอกสารใหม่
      // ทั้งที่ของจริงเป็นเลขวิ่งต่อเนื่องทั้งวันไม่ว่าจะแบ่งบันทึกกี่เอกสารก็ตาม
      const runNoCell = ws.getCell(C.runNo + r);
      runNoCell.value = exportTrackRunningNo(meta.date, (meta.trackStartSeq || 0) + page * cap + i + 1);
      exportSetNumFmt(runNoCell, '@'); // เก็บเป็นข้อความ กันเลข 0 นำหน้าหาย
      ws.getCell(C.po + r).value = it.po;
      ws.getCell(C.supNo + r).value = it.supNo;
      ws.getCell(C.supName + r).value = it.supName;
      exportSetDateCell(ws.getCell(C.recvDate + r), it.recvDateIso);
      ws.getCell(C.confirmDone + r).value = it.confirm === 'done' ? '✓' : null;
      ws.getCell(C.confirmPending + r).value = it.confirm === 'pending' ? '✓' : null;
      ws.getCell(C.remark + r).value = it.remark;
      continue;
    }
    // Barcode เป็นรหัสประจำสินค้า ไม่ใช่ตัวเลขคำนวณ ต้องเก็บเป็นข้อความ + ตั้ง number format
    // เป็น Text (@) เพื่อรักษาเลข 0 นำหน้า และกัน Excel ปัดค่าบาร์โค้ดยาวเป็น scientific notation
    const bcCell = ws.getCell(C.bc + r);
    bcCell.value = String(it.bc == null ? '' : it.bc).trim();
    exportSetNumFmt(bcCell, '@');
    ws.getCell(C.name + r).value = it.name;
    if (mode === 'RTC') {
      ws.getCell(C.unit + r).value = it.unit;
      ws.getCell(C.price + r).value = it.price;
      ws.getCell(C.pct + r).value = it.pct; // decimal (0.2 = 20%) — เซลล์มี numFmt 0% อยู่แล้วในเทมเพลต
      ws.getCell(C.qty + r).value = it.qty;
      // C.newPrice (คอลัมน์ H) ไม่แตะ — ปล่อยให้สูตรเดิมในเทมเพลตคำนวณเองตอนเปิดไฟล์
      exportSetDateCell(ws.getCell(C.exp + r), it.expIso);
    } else {
      ws.getCell(C.code + r).value = it.code;
      ws.getCell(C.qty + r).value = it.qty;
      ws.getCell(C.reason + r).value = it.reason;
      exportSetDateCell(ws.getCell(C.exp + r), it.expIso);
    }
  }

  if (mode === 'RTC') wb.calcProperties.fullCalcOnLoad = true; // บังคับ Excel คำนวณสูตร H ใหม่ตอนเปิดไฟล์

  let buf = await wb.xlsx.writeBuffer();
  buf = await exportFixDrawingXfrm(buf, mode, templateBytes);
  const out = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  // ADJ_<employeeCode>_<date>_<docId>.xlsx
  const filename = mode + '_' + exportSanitizeFilename(meta.employeeCode || 'emp') + '_' +
    exportSanitizeFilename(meta.date || 'form') + '_' + exportSanitizeFilename(meta.docId || mode) +
    (page > 0 ? ('_p' + (page + 1)) : '') + '.xlsx';
  return { blob: out, filename: filename };
};

/*
 * ห่อเอกสารทั้งหมดเป็นไฟล์เดียว ไม่ว่าจะกี่หน้า — เอกสารที่มีรายการเกิน 1 หน้า (เกิน 43/20 แถว)
 * จะถูกรวมเป็นไฟล์ .zip เดียว (มีหลาย .xlsx ข้างในตามจำนวนหน้า) แทนการดาวน์โหลดหลายไฟล์แยกกัน
 * เพราะเทมเพลตต้นฉบับแต่ละหน้าเป็นฟอร์มพิมพ์แยกกันจริง (คนละใบกระดาษ) รวมเป็นชีทเดียวไม่ได้
 * โดยไม่เสี่ยงทำให้ไฟล์ .xlsx เสีย — แต่ยังเป็น "การดาวน์โหลดครั้งเดียว ไฟล์เดียว" ตามที่ต้องการ
 */
// ราคาต่อหน่วยหลังลดราคา — สูตรเดียวกับ calcPrice()/roundThaiCash() ในหน้ามือถือ (ปัดลงทวีคูณ 25 สตางค์
// ตามเครื่องชั่งจริง) คืนค่า null ถ้ารายการนี้ไม่มีราคาต่อหน่วยเดิมบันทึกไว้ (สินค้านับชิ้น ไม่ใช่ชั่งน้ำหนัก)
function exportTagPerUnit(unitPrice, pct) {
  unitPrice = Number(unitPrice); pct = Number(pct);
  if (!unitPrice || !pct) return null;
  return Math.floor(Number((unitPrice - unitPrice * pct).toFixed(6)) * 4) / 4;
}
// ราคาที่ปรับ (ลดเหลือ) — สูตรเดียวกับ roundUpPrice() ในหน้ามือถือ (ปัดขึ้นเป็นจำนวนเต็มบาท)
function exportTagNewPrice(price, pct) {
  price = Number(price); pct = Number(pct);
  if (!price || !pct) return price || 0;
  return Math.ceil(Number((price - price * pct).toFixed(6)));
}
function exportEscHtml(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/*
 * สร้างหน้า HTML สำหรับพิมพ์ป้ายราคาลดราคา (เฉพาะ RTC) แล้วสั่งพิมพ์/บันทึกเป็น PDF ผ่าน print dialog
 * ของเบราว์เซอร์เอง (ไม่มีไลบรารีสร้างไฟล์ .pdf ในเครื่อง — วิธีนี้ได้ PDF จริง เลือกได้ทั้งพิมพ์กระดาษ/
 * Save as PDF โดยไม่ต้องพึ่ง dependency เพิ่ม) ป้ายขนาดจริงดวงละ 5 x 4 ซม. ใช้ CSS Grid บังคับ 4 คอลัมน์
 * แน่นอน (ไม่ใช่ flex-wrap ที่ปัดจำนวนต่อแถวไม่แน่นอนตามพื้นที่ว่างเหลือ) และตัดหน้าทุก 28 ดวง (7 แถว)
 * ด้วย page-break-after ให้ตรง 1 หน้า A4 พอดีเสมอ ไม่ว่าจะมีกี่ดวงก็ตาม
 * แต่ละดวง: บนซ้ายเล็ก = ราคาต่อกก.หลังลดราคา (ถ้ามี — หน่วยคงที่เป็น "กก." เสมอ), กลาง = ชื่อสินค้า + หัวข้อ
 * "ลดเหลือ" ตัวใหญ่ + ราคาที่ปรับแล้วพิมพ์ให้อัตโนมัติ (ตามที่ขอ ไม่เว้นว่างให้เขียนเองอีกต่อไป) + "บาท" —
 * ไม่พิมพ์เลขบาร์โค้ดบนป้าย (ตัดออกตามที่ขอ) แต่ละรายการพิมพ์ซ้ำตามจำนวนชิ้น (qty)
 */
window.buildPriceTagsHtml = function (items) {
  const PER_PAGE = 28; // 4 ดวง/แถว x 7 แถว/หน้า
  const tags = [];
  (items || []).forEach(function (it) {
    const qty = Math.max(1, Number(it.qty) || 1);
    const perKg = exportTagPerUnit(it.unitPrice, it.pct);
    const net = exportTagNewPrice(it.price, it.pct);
    const tagHtml = '<div class="tag">' +
      (perKg != null ? '<div class="perkg">' + perKg.toLocaleString() + ' บ./กก.</div>' : '<div class="perkg">&nbsp;</div>') +
      '<div class="mid"><div class="nm">' + exportEscHtml(it.name) + '</div>' +
      '<div class="reduceLabel">ลดเหลือ</div>' +
      '<div class="netPrice">' + net.toLocaleString() + ' <span class="bahtInline">บาท</span></div></div>' +
      '</div>';
    for (let i = 0; i < qty; i++) tags.push(tagHtml);
  });

  const pages = [];
  for (let i = 0; i < tags.length; i += PER_PAGE) pages.push(tags.slice(i, i + PER_PAGE));
  const pagesHtml = pages.map(function (pageTags, idx) {
    const brk = idx < pages.length - 1 ? ' style="break-after:page;page-break-after:always"' : '';
    return '<div class="sheet"' + brk + '>' + pageTags.join('') + '</div>';
  }).join('');

  return '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>ป้ายราคา</title><style>' +
    '@page{size:A4;margin:8mm}' +
    'body{margin:0;font-family:"Leelawadee UI","Noto Sans Thai",Tahoma,Arial,sans-serif}' +
    '.sheet{display:grid;grid-template-columns:repeat(4,5cm);grid-auto-rows:4cm;justify-content:center}' +
    '.tag{width:5cm;height:4cm;box-sizing:border-box;border:1px solid #999;' +
    'padding:2mm;display:flex;flex-direction:column;overflow:hidden;break-inside:avoid;page-break-inside:avoid}' +
    '.perkg{font-size:9px;color:#555}' +
    '.mid{text-align:center;margin-top:2mm}' +
    '.nm{font-size:11px;font-weight:700;line-height:1.2;display:-webkit-box;-webkit-line-clamp:1;' +
    '-webkit-box-orient:vertical;overflow:hidden}' +
    '.reduceLabel{font-size:15px;font-weight:800;margin-top:1mm}' +
    '.netPrice{font-size:22px;font-weight:800;margin-top:1mm}' +
    '.bahtInline{font-size:12px;font-weight:600}' +
    '.noprint{padding:10px}' +
    '@media print{.noprint{display:none}}' +
    '</style></head><body>' +
    '<div class="noprint"><button onclick="window.print()">พิมพ์ / บันทึกเป็น PDF</button></div>' +
    pagesHtml +
    '</body></html>';
};

window.buildExportPackage = async function (mode, items, meta) {
  const cfg = EXPORT_CFG[mode], cap = cfg.last - cfg.first + 1;
  const pages = Math.max(1, Math.ceil((items || []).length / cap));
  if (pages <= 1) {
    return window.buildExportXlsx(mode, items, meta, 0);
  }
  const pkg = new JSZip();
  for (let p = 0; p < pages; p++) {
    const { blob, filename } = await window.buildExportXlsx(mode, items, meta, p);
    pkg.file(filename, blob);
  }
  const out = await pkg.generateAsync({ type: 'blob', mimeType: 'application/zip' });
  const filename = mode + '_' + exportSanitizeFilename(meta.employeeCode || 'emp') + '_' +
    exportSanitizeFilename(meta.date || 'form') + '_' + exportSanitizeFilename(meta.docId || mode) +
    '_' + pages + 'pages.zip';
  return { blob: out, filename: filename };
};
