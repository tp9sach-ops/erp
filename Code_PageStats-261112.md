/** ============================================================
 *  9 SẠCH - THỐNG KÊ TRANG (KHÁCH MỚI / ĐƠN CHỐT)  v1.0
 *  ================================================================
 *  CÁCH DÙNG: Trong Apps Script Editor, bấm dấu "+" cạnh "Files" ->
 *  Script -> đặt tên "Code_PageStats" -> dán TOÀN BỘ nội dung file
 *  này vào. Không cần sửa Code.gs - dùng lại _sheet() và _clean()
 *  đã có sẵn.
 *
 *  Đọc đúng cấu trúc file "Thống kê trang" (Thời gian, Khách mới,
 *  SĐT trong ngày, SĐT mới, Số đơn chốt...) - mỗi dòng = 1 NGÀY.
 *  Upsert theo Ngày -> KHÔNG BAO GIỜ tạo trùng ngày dù upload lại
 *  nhiều lần. Có thêm dedupePageStatsSheet() để dọn dữ liệu cũ nếu
 *  đã lỡ bị trùng từ trước khi có cơ chế này.
 * ============================================================ */

const PS_DAILY_HEADERS = ['Ngày', 'Khách Mới', 'SĐT Trong Ngày', 'SĐT Mới', 'Số Đơn Chốt', 'Tỷ Lệ Đơn/KH Mới (%)', 'Cập Nhật Lúc', 'Batch ID Cuối'];
function _psDailySheet() {
  return _sheet('ThongKeTrangNgay', PS_DAILY_HEADERS);
}
const PS_LOG_HEADERS = ['Batch ID', 'Thời Gian Upload', 'Tên File', 'Số Ngày Thêm Mới', 'Số Ngày Cập Nhật', 'Nhân Viên Upload'];
function _psUploadLogSheet() { return _sheet('ThongKeTrang_UploadLog', PS_LOG_HEADERS); }

function _psDayKey(dateVal) {
  const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return '';
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}

// ════════════════════════════════════════════════════════════
// NHẬP 1 LÔ NGÀY (upload file)
// rowsArr: [{dateStr, khachMoi, sdtNgay, sdtMoi, donChot}, ...]
// - Dedupe NGAY TRONG LÔ đang upload: nếu file có 2 dòng cùng ngày,
//   chỉ giữ dòng cuối.
// - Dedupe VỚI DỮ LIỆU ĐÃ CÓ: ngày đã tồn tại trong sheet -> GHI ĐÈ,
//   không thêm dòng mới -> không bao giờ tạo trùng ngày.
// ════════════════════════════════════════════════════════════
function importPageStatsBatch(rowsArr, fileName, employeeName) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    const byDay = {};
    (rowsArr || []).forEach(function (r) {
      const key = _psDayKey(r.dateStr);
      if (key) byDay[key] = r; // dòng cuối trong file thắng nếu trùng ngày
    });

    const sh = _psDailySheet();
    const batchId = 'PS' + new Date().getTime();
    const lastRow = sh.getLastRow();
    const existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, PS_DAILY_HEADERS.length).getValues() : [];
    const dayIdx = {};
    existing.forEach(function (r, i) {
      const key = _psDayKey(r[0]);
      if (key) dayIdx[key] = i + 2;
    });

    let added = 0, updated = 0;
    const newRows = [];
    Object.keys(byDay).forEach(function (key) {
      const r = byDay[key];
      const khachMoi = Number(r.khachMoi) || 0;
      const donChot = Number(r.donChot) || 0;
      const ratio = khachMoi > 0 ? (donChot / khachMoi * 100) : 0;
      const rowArr = [new Date(r.dateStr), khachMoi, Number(r.sdtNgay) || 0, Number(r.sdtMoi) || 0, donChot,
      Math.round(ratio * 100) / 100, new Date(), batchId];
      if (dayIdx[key]) {
        sh.getRange(dayIdx[key], 1, 1, rowArr.length).setValues([rowArr]);
        updated++;
      } else {
        newRows.push(rowArr);
        added++;
      }
    });
    if (newRows.length) {
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, newRows.length, PS_DAILY_HEADERS.length).setValues(newRows);
    }

    const logSh = _psUploadLogSheet();
    logSh.appendRow([batchId, new Date(), fileName || '', added, updated, employeeName || '']);

    return { success: true, added: added, updated: updated, batchId: batchId };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally { lock.releaseLock(); }
}

function getPageStatsUploadLog() {
  try {
    const sh = _psUploadLogSheet();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];
    const data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    return data.map(function (r) {
      return { batchId: (r[0] || '').toString(), time: _clean(r[1]), fileName: r[2] || '', added: Number(r[3]) || 0, updated: Number(r[4]) || 0, emp: r[5] || '' };
    }).sort(function (a, b) { return new Date(b.time) - new Date(a.time); });
  } catch (e) { return []; }
}

// Xóa 1 lượt upload - chỉ xóa các ngày mà "Batch ID Cuối" còn đúng bằng batchId
// này (nếu ngày đó đã bị 1 lượt upload SAU ghi đè thì sẽ không bị xóa nhầm).
function deletePageStatsUploadBatch(batchId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    const sh = _psDailySheet();
    const lastRow = sh.getLastRow();
    let deletedCount = 0;
    if (lastRow > 1) {
      const data = sh.getRange(2, 1, lastRow - 1, PS_DAILY_HEADERS.length).getValues();
      const rowsToDelete = [];
      data.forEach(function (r, i) { if ((r[7] || '').toString() === batchId) rowsToDelete.push(i + 2); });
      rowsToDelete.sort(function (a, b) { return b - a; });
      rowsToDelete.forEach(function (rowIdx) { sh.deleteRow(rowIdx); deletedCount++; });
    }
    const logSh = _psUploadLogSheet();
    const logLastRow = logSh.getLastRow();
    if (logLastRow > 1) {
      const logData = logSh.getRange(2, 1, logLastRow - 1, 1).getValues();
      for (let i = 0; i < logData.length; i++) {
        if ((logData[i][0] || '').toString() === batchId) { logSh.deleteRow(i + 2); break; }
      }
    }
    return { success: true, deletedCount: deletedCount };
  } catch (e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

// [TÍNH NĂNG XÓA TRÙNG] Dọn dữ liệu trùng NGÀY trong toàn bộ sheet (phòng khi
// dữ liệu cũ đã lỡ bị lưu trùng trước khi có cơ chế upsert). Giữ dòng có
// "Cập Nhật Lúc" mới nhất cho mỗi ngày, xóa hết các dòng trùng còn lại.
function dedupePageStatsSheet() {
  try {
    const sh = _psDailySheet();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, removed: 0 };
    const data = sh.getRange(2, 1, lastRow - 1, PS_DAILY_HEADERS.length).getValues();
    const bestByDay = {};
    data.forEach(function (r) {
      const key = _psDayKey(r[0]);
      if (!key) return;
      const updatedAt = r[6] instanceof Date ? r[6].getTime() : 0;
      if (!bestByDay[key] || updatedAt >= bestByDay[key].updatedAt) {
        bestByDay[key] = { updatedAt: updatedAt, row: r.slice() };
      }
    });
    const keptRows = Object.keys(bestByDay).sort().map(function (k) { return bestByDay[k].row; });
    const removed = data.length - keptRows.length;
    sh.getRange(2, 1, data.length, PS_DAILY_HEADERS.length).clearContent();
    if (keptRows.length) {
      sh.getRange(2, 1, keptRows.length, PS_DAILY_HEADERS.length).setValues(keptRows);
    }
    return { success: true, removed: removed };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function _psLoadAll() {
  const sh = _psDailySheet();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, PS_DAILY_HEADERS.length).getValues().map(function (r) {
    return {
      date: r[0] instanceof Date ? r[0] : new Date(r[0]),
      khachMoi: Number(r[1]) || 0, sdtNgay: Number(r[2]) || 0, sdtMoi: Number(r[3]) || 0,
      donChot: Number(r[4]) || 0, ratio: Number(r[5]) || 0
    };
  });
}

function _psMonthAgg(rows, y, m) {
  let khachMoi = 0, donChot = 0;
  rows.forEach(function (r) {
    if (r.date.getFullYear() === y && r.date.getMonth() === m) { khachMoi += r.khachMoi; donChot += r.donChot; }
  });
  const ratio = khachMoi > 0 ? (donChot / khachMoi * 100) : 0;
  return { khachMoi: khachMoi, donChot: donChot, ratio: ratio };
}

// [FIX] Hàm này bị thiếu, gây lỗi "_psTotalsForMonth is not defined" ở
// trang Thưởng KPI và khi Xuất Văn Bản Target. Tính tổng Khách Mới,
// SĐT Trong Ngày, SĐT Mới, Đơn Chốt từ sheet ThongKeTrangNgay cho
// đúng 1 tháng cụ thể - dùng để tính tỷ lệ chốt đơn so với SĐT mới.
function _psTotalsForMonth(monthKey) {
  try {
    const parts = (monthKey || '').split('-');
    const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;
    const rows = _psLoadAll();
    let khachMoi = 0, sdtNgay = 0, sdtMoi = 0, donChot = 0;
    rows.forEach(function (r) {
      if (r.date.getFullYear() === y && r.date.getMonth() === m) {
        khachMoi += r.khachMoi;
        sdtNgay += r.sdtNgay;
        sdtMoi += r.sdtMoi;
        donChot += r.donChot;
      }
    });
    const ratio = khachMoi > 0 ? (donChot / khachMoi * 100) : 0;
    return { khachMoi: khachMoi, sdtNgay: sdtNgay, sdtMoi: sdtMoi, donChot: donChot, ratio: ratio };
  } catch (e) {
    return { khachMoi: 0, sdtNgay: 0, sdtMoi: 0, donChot: 0, ratio: 0 };
  }
}

// Bảng theo năm: Khách Mới / Số Đơn Chốt / Tỷ Lệ Đơn-Khách Mới theo 12 tháng,
// kèm Target = TB tỷ lệ của 3 tháng liền trước mỗi tháng (tự động lùi qua năm
// trước nếu tháng đang xét là T1/T2/T3).
function getPageStatsYearTable(year) {
  try {
    const y = parseInt(year, 10) || new Date().getFullYear();
    const rows = _psLoadAll();
    const khachMoiArr = [], donChotArr = [], ratioArr = [], targetArr = [];
    for (let m = 0; m < 12; m++) {
      const agg = _psMonthAgg(rows, y, m);
      khachMoiArr.push(agg.khachMoi); donChotArr.push(agg.donChot); ratioArr.push(Math.round(agg.ratio * 100) / 100);
      let sum = 0, cnt = 0;
      for (let i = 1; i <= 3; i++) {
        let pm = m - i, py = y; while (pm < 0) { pm += 12; py--; }
        const pAgg = _psMonthAgg(rows, py, pm);
        if (pAgg.khachMoi > 0) { sum += pAgg.ratio; cnt++; }
      }
      targetArr.push(cnt ? Math.round(sum / cnt * 100) / 100 : 0);
    }
    return { year: y, khachMoi: khachMoiArr, donChot: donChotArr, ratio: ratioArr, target: targetArr };
  } catch (e) { return { error: e.toString() }; }
}

// Chi tiết theo ngày của 1 tháng - dùng khi bấm vào số trên bảng tháng.
function getPageStatsMonthDetail(year, month) {
  try {
    const y = parseInt(year, 10), m = parseInt(month, 10);
    const rows = _psLoadAll().filter(function (r) { return r.date.getFullYear() === y && r.date.getMonth() === m; });
    rows.sort(function (a, b) { return a.date - b.date; });
    return rows.map(function (r) {
      return { date: r.date.toISOString(), khachMoi: r.khachMoi, sdtNgay: r.sdtNgay, sdtMoi: r.sdtMoi, donChot: r.donChot, ratio: r.ratio };
    });
  } catch (e) { return []; }
}