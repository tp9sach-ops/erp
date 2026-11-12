/** ============================================================
 *  9 SẠCH - MODULE KÊNH MT (AEON / LOTTE / SATRA / WINMART / KINGKONG)  v1.1
 *  ================================================================
 *  CÁCH DÙNG: dán ĐÈ toàn bộ file "MT_Channel_Backend" hiện có trong
 *  Apps Script bằng nội dung này. Không cần sửa Code.gs - các hàm
 *  _sheet(), _clean(), _monthKey() đã có sẵn sẽ được dùng lại.
 *  v1.1: thêm kênh KINGKONG (danh sách kênh tập trung ở MT_CHANNELS).
 * ============================================================ */

// Danh sách kênh MT - thêm kênh mới chỉ cần sửa đúng 1 chỗ này
const MT_CHANNELS = ['AEON', 'LOTTE', 'SATRA', 'WINMART', 'KINGKONG', 'CUONGTHANH'];

const MT_STATS_HEADERS = ['Ngày Đặt', 'Kênh MT', 'Số PO/Đơn Hàng', 'Khách/Siêu Thị Nhận',
  'Số Thùng', 'Doanh Thu Chưa Thuế', 'VAT', 'Tổng Thanh Toán', 'Cập Nhật Lúc', 'Nhân Viên',
  'Người Duyệt', 'Thời Gian Duyệt', 'Cần Duyệt'];

function _mtStatsSheet() {
  const sh = _sheet('MT_Stats', MT_STATS_HEADERS);
  try {
    const hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 13)).getValues()[0];
    if (hdr[10] !== 'Người Duyệt') sh.getRange(1, 11, 1, 3).setValues([['Người Duyệt', 'Thời Gian Duyệt', 'Cần Duyệt']]);
  } catch (e) {}
  // [MỚI] Cột C (Số PO) và D (Siêu thị) luôn là VĂN BẢN: tránh Sheets đổi "0012345" thành 12345 hay "3/2" thành ngày
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('MT_TEXTFMT_V1') !== '1') {
      sh.getRange(2, 3, Math.max(sh.getMaxRows() - 1, 1), 2).setNumberFormat('@');
      props.setProperty('MT_TEXTFMT_V1', '1');
    }
  } catch (e) {}
  return sh;
}
// Dòng được tính doanh thu: dòng cũ (cột "Cần Duyệt" trống) hoặc dòng mới đã có người duyệt
function _mtRowCounted(r) {
  const needs = (r[12] || '').toString() === 'Y';
  return !needs || !!(r[10] || '').toString().trim();
}

function saveMTOrderStat(channel, poNo, storeName, boxCount, subTotal, vatAmt, grandTotal, dateStr, employeeName) {
  try {
    channel = (channel || '').toString().trim().toUpperCase();
    if (MT_CHANNELS.indexOf(channel) === -1) return { success: false, error: 'Kênh MT "' + channel + '" không hợp lệ.' };
    const sh = _mtStatsSheet();
    const lastRow = sh.getLastRow();
    const keys = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 3).getValues() : [];
    const newDate = new Date(dateStr || new Date());
    const newMonthKey = _monthKey(newDate);
    const row = [
      newDate, channel, (poNo || '').toString(), storeName || '',
      Number(boxCount) || 0, Number(subTotal) || 0, Number(vatAmt) || 0, Number(grandTotal) || 0,
      new Date(), employeeName || '', '', '', 'Y'
    ];
    for (let i = 0; i < keys.length; i++) {
      if ((keys[i][1] || '') === channel && (keys[i][2] || '').toString() === (poNo || '').toString()) {
        const oldDate = keys[i][0] instanceof Date ? keys[i][0] : new Date(keys[i][0]);
        const oldMonthKey = _monthKey(oldDate);
        if (oldMonthKey !== newMonthKey) {
          return {
            success: false,
            error: 'Đơn "' + poNo + '" đã lưu ở tháng ' + oldMonthKey + ' trước đó. ' +
              'Muốn dời sang tháng ' + newMonthKey + ', hãy vào "Xem đơn hàng kênh MT" xóa dòng cũ rồi lưu lại.'
          };
        }
        const oldRowFull = sh.getRange(i + 2, 1, 1, 13).getValues()[0];
        const oldNeeds = (oldRowFull[12] || '').toString() === 'Y';
        const oldApprovedBy = (oldRowFull[10] || '').toString().trim();
        row[9] = (oldRowFull[9] || '').toString().trim() || row[9];
        if (!oldNeeds) {
          row[12] = '';
        } else if (oldApprovedBy && (Number(oldRowFull[7]) || 0) === (Number(grandTotal) || 0)) {
          row[10] = oldApprovedBy; row[11] = oldRowFull[11];
        }
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return { success: true, updated: true };
      }
    }
    sh.appendRow(row);
    return { success: true, updated: false };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// [MỚI] true = MỌI đơn MT mới đều phải duyệt (như cũ). false = đơn có ngày trước 01/10/2026 tự được tính (giống DonHang).
const MT_ALWAYS_NEED_APPROVAL = false;

// [MỚI] Chuẩn hóa số PO để so sánh: bỏ khoảng trắng, bỏ số 0 đầu của PO toàn chữ số
function _mtPoKey_(v) {
  let s = (v === null || v === undefined) ? '' : v.toString().trim();
  if (/^\d+$/.test(s)) s = s.replace(/^0+(?=\d)/, '');
  return s;
}
// [MỚI] Ô bất kỳ -> chuỗi an toàn để gửi về trình duyệt (Date không gửi được qua google.script.run)
function _mtStr_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  return (v === null || v === undefined) ? '' : v.toString();
}

// [SỬA] Đọc sheet 1 lần, khóa, ghi theo lô. Trước đây mỗi đơn đọc lại cả cột A:C và ghi riêng từng dòng
// (300 đơn x sheet vài nghìn dòng = rất chậm, dễ quá 6 phút), không có khóa nên 2 người lưu cùng lúc có thể ra đơn trùng.
function saveMTOrderStatsBatch(rowsArr) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, vui lòng thử lại sau ít phút.' }; }
  try {
    const sh = _mtStatsSheet();
    const lastRow = sh.getLastRow();
    const all = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 13).getValues() : [];
    const origLen = all.length;
    const index = {};
    all.forEach(function (r, i) {
      if (!r[1] && !r[2]) return;
      const k = (r[1] || '').toString().trim().toUpperCase() + '|' + _mtPoKey_(r[2]);
      if (!(k in index)) index[k] = i;
    });
    const errors = [];
    const changed = {};
    let ok = 0, added = 0, updated = 0, unchanged = 0, pending = 0;
    (rowsArr || []).forEach(function (r) {
      const channel = (r.channel || '').toString().trim().toUpperCase();
      if (MT_CHANNELS.indexOf(channel) === -1) { errors.push('Kênh MT "' + channel + '" không hợp lệ.'); return; }
      const poNo = (r.poNo || '').toString().trim();
      if (!poNo) { errors.push('Có đơn thiếu số PO (' + (r.storeName || 'không rõ siêu thị') + ') - bỏ qua.'); return; }
      const newDate = new Date(r.dateStr || new Date());
      if (isNaN(newDate.getTime())) { errors.push('Đơn "' + poNo + '": ngày không hợp lệ (' + r.dateStr + ').'); return; }
      const grand = Number(r.grandTotal) || 0;
      const needs = MT_ALWAYS_NEED_APPROVAL || _orderNeedsApproval(newDate);
      const row = [newDate, channel, poNo, r.storeName || '', Number(r.boxCount) || 0, Number(r.subTotal) || 0, Number(r.vatAmt) || 0, grand,
        new Date(), r.employeeName || '', '', '', needs ? 'Y' : ''];
      const key = channel + '|' + _mtPoKey_(poNo);
      if (key in index) {
        const i = index[key], old = all[i];
        const oldDate = old[0] instanceof Date ? old[0] : new Date(old[0]);
        if (_monthKey(oldDate) !== _monthKey(newDate)) {
          errors.push('Đơn "' + poNo + '" đã lưu ở tháng ' + _monthKey(oldDate) + ' trước đó. Muốn dời sang tháng ' + _monthKey(newDate) +
            ', hãy vào "Xem đơn hàng kênh MT" xóa dòng cũ rồi lưu lại.');
          return;
        }
        const oldNeeds = (old[12] || '').toString() === 'Y';
        const oldBy = (old[10] || '').toString().trim();
        // [GIỮ GỐC] upload lại: giữ nhân viên đã nhập đơn lần đầu (chỉ lấy tên mới nếu ô cũ trống)
        row[9] = (old[9] || '').toString().trim() || row[9];
        const same = _mtStr_(oldDate) === _mtStr_(newDate) && (old[3] || '').toString() === row[3] &&
          (Number(old[4]) || 0) === row[4] && Math.round(Number(old[5]) || 0) === Math.round(row[5]) &&
          Math.round(Number(old[6]) || 0) === Math.round(row[6]) && Math.round(Number(old[7]) || 0) === Math.round(grand);
        ok++;
        if (same) { unchanged++; if (oldNeeds && !oldBy) pending++; return; }
        if (!oldNeeds) row[12] = '';
        else if (oldBy && Math.round(Number(old[7]) || 0) === Math.round(grand)) { row[10] = oldBy; row[11] = old[11]; }
        all[i] = row;
        if (i < origLen) changed[i] = true;
        updated++;
        if (row[12] === 'Y' && !row[10]) pending++;
      } else {
        all.push(row);
        index[key] = all.length - 1;
        ok++; added++;
        if (row[12] === 'Y') pending++;
      }
    });
    Object.keys(changed).forEach(function (k) {
      const i = Number(k);
      sh.getRange(i + 2, 1, 1, 13).setValues([all[i]]);
    });
    if (all.length > origLen) {
      const fresh = all.slice(origLen);
      sh.getRange(lastRow + 1, 1, fresh.length, 13).setValues(fresh);
    }
    if ((added || updated) && typeof _kpiBust_ === 'function') { try { _kpiBust_(); } catch (e1) {} }
    return { success: true, count: ok, errors: errors, added: added, updated: updated, unchanged: unchanged, pending: pending };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally { try { lock.releaseLock(); } catch (x) {} }
}

function getMTStats(fromDate, toDate) {
  try {
    const sh = _mtStatsSheet();
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    const from = fromDate ? new Date(fromDate) : null;
    const to = toDate ? new Date(toDate) : null;
    if (to) to.setHours(23, 59, 59, 999);
    return data.slice(1).filter(function (r) {
      if (!r[0]) return false;
      if (!_mtRowCounted(r)) return false;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    }).map(function (r) {
      return {
        date: _clean(r[0]), channel: r[1] || '', poNo: (r[2] || '').toString(), store: _mtStr_(r[3]),
        boxes: Number(r[4]) || 0, subTotal: Number(r[5]) || 0, vat: Number(r[6]) || 0,
        total: Number(r[7]) || 0, emp: r[9] || ''
      };
    });
  } catch (e) { Logger.log('getMTStats lỗi: ' + e); return []; }
}

// Tổng hợp theo từng kênh MT - thẻ + biểu đồ trong "Sức Khỏe Kinh Doanh". Mặc định THÁNG NÀY.
function getMTChannelSummary(fromDate, toDate) {
  try {
    let from = fromDate, to = toDate;
    if (!from && !to) {
      const now = new Date();
      from = new Date(now.getFullYear(), now.getMonth(), 1).toISOString();
      to = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59).toISOString();
    }
    const rows = getMTStats(from, to);
    const channels = MT_CHANNELS;
    const byChannel = {};
    channels.forEach(function (c) { byChannel[c] = { channel: c, orders: 0, boxes: 0, subTotal: 0, vat: 0, total: 0 }; });
    rows.forEach(function (r) {
      const c = (r.channel || '').toUpperCase();
      if (!byChannel[c]) byChannel[c] = { channel: c, orders: 0, boxes: 0, subTotal: 0, vat: 0, total: 0 };
      byChannel[c].orders++;
      byChannel[c].boxes += r.boxes;
      byChannel[c].subTotal += r.subTotal;
      byChannel[c].vat += r.vat;
      byChannel[c].total += r.total;
    });
    const list = Object.keys(byChannel);
    const grandTotal = list.reduce(function (a, c) { return a + byChannel[c].total; }, 0);
    list.forEach(function (c) { byChannel[c].pct = grandTotal ? (byChannel[c].total / grandTotal * 100) : 0; });
    return { byChannel: list.map(function (c) { return byChannel[c]; }), grandTotal: grandTotal, from: from, to: to };
  } catch (e) {
    return { byChannel: [], grandTotal: 0, error: e.toString() };
  }
}

// ════════════════════════════════════════════════════════════
//  KÊNH AEON - PARSE FILE CSV (giữ nguyên v1.0)
// ════════════════════════════════════════════════════════════
function parseCSVLine(line) {
  const out = []; let cur = ''; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === ',' && !inQ) { out.push(cur); cur = ''; }
    else cur += ch;
  }
  out.push(cur);
  return out.map(function (s) { return s.replace(/^"|"$/g, '').trim(); });
}

function formatDateStr(dateStr) {
  const s = (dateStr || '').toString().trim();
  if (/^\d{8}$/.test(s)) return s.slice(6, 8) + '/' + s.slice(4, 6) + '/' + s.slice(0, 4);
  return s;
}

function processAeonCSVFiles(fileDataArray) {
  try {
    const allOrders = {};
    const seenLine = {};   // [MỚI] 'PO|dòng gốc' -> số thứ tự file đã đọc dòng này đầu tiên
    (fileDataArray || []).forEach(function (fileText, fileIdx) {
      if (!fileText) return;
      const lines = fileText.split(/\r?\n/);
      let headers = null;
      lines.forEach(function (line) {
        if (!line || !line.trim()) return;
        const trimmed = line.trim().replace(/^"|"$/g, '');
        if (/^Bat dau$/i.test(trimmed) || /^Ket thuc$/i.test(trimmed)) return;
        const cols = parseCSVLine(line);
        if (!headers) {
          if (cols[0] && cols[0].toLowerCase().indexOf('loai nganh hang') !== -1) { headers = cols; }
          return;
        }
        if (cols.length < 10 || !cols[1]) return;
        const rec = {};
        headers.forEach(function (h, idx) { rec[h] = cols[idx]; });
        const poNo = rec['So Don hang'];
        if (!poNo) return;
        // [MỚI] Cùng 1 dòng hàng của cùng 1 PO xuất hiện ở file khác (2 file chọn chồng ngày) -> bỏ, không cộng đôi tiền
        const lineKey = poNo + '|' + line.trim();
        if (lineKey in seenLine && seenLine[lineKey] !== fileIdx) return;
        seenLine[lineKey] = fileIdx;
        if (!allOrders[poNo]) {
          allOrders[poNo] = {
            poNo: poNo, storeCode: rec['Ma cua hang'] || '', storeName: rec['Ten cua hang'] || '',
            orderDate: formatDateStr(rec['Ngay dat hang']), deliveryDate: formatDateStr(rec['Ngay giao hang']),
            vendorCode: rec['Ma NCC'] || '', contractNo: rec['So Hop dong'] || '',
            deliveryPlace: rec['Dia diem giao hang'] || '',
            items: [], subTotal: 0, vatAmt: 0, total: 0
          };
        }
        const qty = parseFloat(rec['So luong']) || 0;
        const unitPrice = parseFloat(rec['Don gia mua chua thue VAT']) || 0;
        allOrders[poNo].items.push({
          name: rec['Ten hang'] || '', code: rec['Ma hang'] || '', barcode: rec['Ma vach'] || '',
          unit: rec['DON VI TINH'] || '', qty: qty, unitPrice: unitPrice, lineTotal: qty * unitPrice
        });
      });
    });
    const result = Object.keys(allOrders).map(function (k) {
      const o = allOrders[k];
      o.subTotal = o.items.reduce(function (a, it) { return a + it.lineTotal; }, 0);
      o.vatAmt = Math.round(o.subTotal * 0.08);
      o.total = o.subTotal + o.vatAmt;
      return o;
    });
    return result;
  } catch (e) {
    return { error: e.toString() };
  }
}
// ════════════════════════════════════════════════════════════
//  [MỚI] KIỂM TRA SHEET MT_Stats (CHỈ ĐỌC) - chạy tay, dán nhật ký cho người hỗ trợ
//  Báo: tổng theo tháng/kênh (đã tính vs chờ duyệt), PO trùng, ngày lỗi/ngày tương lai,
//  ô PO/Siêu thị sai kiểu, dòng có Tổng ≠ Chưa thuế + VAT.
// ════════════════════════════════════════════════════════════
function kiemTraMTStats() {
  const sh = _mtStatsSheet();
  const last = sh.getLastRow();
  const L = ['===== KIỂM TRA MT_Stats (chỉ đọc) - ' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss') + ' ====='];
  if (last < 2) { L.push('Sheet trống.'); console.log(L.join('\n')); return; }
  const vals = sh.getRange(2, 1, last - 1, 13).getValues();
  const now = new Date();
  const byM = {}, seen = {}, badDate = [], future = [], dateInStore = [], numPo = [], lostZero = [], sumBad = [], noEmp = [];
  vals.forEach(function (r, i) {
    const row = i + 2;
    if (!r[0] && !r[1] && !r[2]) return;
    const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
    if (isNaN(d.getTime())) { badDate.push(row); return; }
    if (d.getTime() > now.getTime() + 2 * 86400000) future.push(row + ' (' + Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') + ')');
    const mk = Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'yyyy-MM');
    const ch = (r[1] || '(trống)').toString();
    const m = byM[mk] || (byM[mk] = {});
    const c = m[ch] || (m[ch] = { n: 0, counted: 0, pending: 0, rev: 0, pendRev: 0 });
    const tot = Number(r[7]) || 0;
    c.n++;
    if (_mtRowCounted(r)) { c.counted++; c.rev += tot; } else { c.pending++; c.pendRev += tot; }
    const k = ch.toUpperCase() + '|' + _mtPoKey_(r[2]);
    (seen[k] = seen[k] || []).push(row);
    if (r[3] instanceof Date) dateInStore.push(row);
    if (typeof r[2] === 'number') { numPo.push(row); if (r[2] >= 1e15) lostZero.push(row); }
    if (Math.abs((Number(r[5]) || 0) + (Number(r[6]) || 0) - tot) > 1) sumBad.push(row);
    if (!(r[9] || '').toString().trim()) noEmp.push(row);
  });
  const fmt = function (n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.') + 'đ'; };
  Object.keys(byM).sort().forEach(function (mk) {
    L.push('--- Tháng ' + mk);
    Object.keys(byM[mk]).sort().forEach(function (ch) {
      const c = byM[mk][ch];
      L.push('  ' + ch + ': ' + c.n + ' đơn | đã tính ' + c.counted + ' = ' + fmt(c.rev) + (c.pending ? ' | CHỜ DUYỆT ' + c.pending + ' đơn = ' + fmt(c.pendRev) + ' (chưa tính doanh thu)' : ''));
    });
  });
  const dups = Object.keys(seen).filter(function (k) { return seen[k].length > 1; });
  L.push('', 'PO TRÙNG (cùng kênh + cùng số PO): ' + (dups.length ? dups.length : 'không'));
  dups.slice(0, 15).forEach(function (k) { L.push('  ' + k + ' -> dòng ' + seen[k].join(', ')); });
  L.push('Ngày không đọc được: ' + (badDate.length ? 'dòng ' + badDate.slice(0, 20).join(', ') : 'không'));
  L.push('Ngày ở tương lai (> 2 ngày): ' + (future.length ? future.slice(0, 20).join(', ') : 'không'));
  L.push('Ô Siêu thị bị Sheets đổi thành ngày: ' + (dateInStore.length ? 'dòng ' + dateInStore.slice(0, 20).join(', ') : 'không'));
  L.push('Ô Số PO bị lưu thành SỐ: ' + (numPo.length ? numPo.length + ' dòng (nguy cơ mất số 0 đầu)' : 'không') + (lostZero.length ? ' | quá 15 chữ số, đã mất độ chính xác: dòng ' + lostZero.slice(0, 10).join(', ') : ''));
  L.push('Dòng có Tổng ≠ Chưa thuế + VAT (lệch > 1đ): ' + (sumBad.length ? 'dòng ' + sumBad.slice(0, 20).join(', ') : 'không'));
  L.push('Dòng thiếu tên nhân viên: ' + (noEmp.length ? noEmp.length + ' dòng' : 'không'));
  L.push('Hàm này KHÔNG ghi gì.');
  console.log(L.join('\n'));
}