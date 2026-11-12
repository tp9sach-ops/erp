/** ============================================================
 *  9 SẠCH - DUYỆT ĐƠN HÀNG TẠO (áp dụng từ 01/10/2026)
 *  ================================================================
 *  CÁCH DÙNG: Apps Script Editor -> "+" cạnh Files -> Script -> đặt tên
 *  "Code_DuyetDon" -> dán TOÀN BỘ nội dung này.
 *
 *  QUY TẮC: Đơn có ngày đặt từ 01/10/2026 chỉ được GHI NHẬN doanh thu /
 *  số lượng / KPI sau khi được duyệt. Đơn trước 01/10/2026 tính bình thường.
 *  Tạo đơn, sao chép, tra cứu, lịch sử... KHÔNG thay đổi.
 *  Người duyệt + thời gian duyệt lưu ở cột X, Y của sheet DonHang
 *  (chỉ dòng đầu của mỗi đơn).
 * ============================================================ */
const APPROVAL_START = new Date(2026, 9, 1, 0, 0, 0, 0);   // 01/10/2026
const APPROVER_ROLES = ['Admin', 'QLB2B'];                   // quyền được bấm Duyệt - sửa tại đây nếu cần
const COL_APPROVER = 24;
const COL_APPROVED_AT = 25;

function _orderNeedsApproval(d) {
  const dt = d instanceof Date ? d : new Date(d);
  return !isNaN(dt.getTime()) && dt.getTime() >= APPROVAL_START.getTime();
}
// Khối đơn bắt đầu tại dòng "row" có được tính doanh thu/KPI chưa?
function _isBlockCounted(sh, row) {
  const v = sh.getRange(row, 1, 1, COL_APPROVED_AT).getValues()[0];
  if (!_orderNeedsApproval(v[0])) return true;
  return _approvalValid_(v[COL_APPROVER - 1], v[COL_APPROVED_AT - 1]);
}
function _apprDayKey(d) {
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}

// Tổng hợp theo từng ngày của 1 tháng (để chọn ngày)
function getApprovalDaySummary(monthKey) {
  try {
    const parts = (monthKey || '').split('-');
    if (parts.length !== 2) return { error: 'Tháng không hợp lệ' };
    const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;
    const from = new Date(y, m, 1, 0, 0, 0, 0), to = new Date(y, m + 1, 0, 23, 59, 59, 999);
    const orders = _readDonHangGrouped(from.toISOString(), to.toISOString(), null);
    const byDay = {};
    orders.forEach(function (o) {
      const k = _apprDayKey(new Date(o.date));
      if (!byDay[k]) byDay[k] = { day: k, orders: 0, pending: 0, approved: 0, notRequired: 0, revenue: 0 };
      const d = byDay[k];
      d.orders++;
      if (!o.needsApproval) d.notRequired++;
      else if (o.approvedBy) { d.approved++; d.revenue += o.total; }
      else d.pending++;
      if (!o.needsApproval) d.revenue += o.total;
    });
    const days = Object.keys(byDay).sort().map(function (k) { return byDay[k]; });
    return { monthKey: monthKey, days: days, startDate: '2026-10-01' };
  } catch (e) { return { error: e.toString() }; }
}

// Danh sách đơn của 1 ngày (dayStr = 'yyyy-mm-dd') kèm chi tiết sản phẩm
function getOrdersForApprovalDay(dayStr) {
  try {
    const p = (dayStr || '').split('-');
    if (p.length !== 3) return { error: 'Ngày không hợp lệ' };
    const y = parseInt(p[0], 10), m = parseInt(p[1], 10) - 1, d = parseInt(p[2], 10);
    const from = new Date(y, m, d, 0, 0, 0, 0), to = new Date(y, m, d, 23, 59, 59, 999);
    const orders = _readDonHangGrouped(from.toISOString(), to.toISOString(), null);
    const list = orders.map(function (o) {
      let qty = 0;
      o.items.forEach(function (it) { qty += Number(it.qty) || 0; });
      return {
        sheetStartRow: o.sheetStartRow, rowCount: o.items.length, date: o.date,
        name: o.name, phone: o.phone, emp: o.emp, qty: qty, total: o.total,
        priceType: o.priceType, note: o.note, address: o.address,
        needsApproval: !!o.needsApproval, approvedBy: o.approvedBy || '', approvedAt: o.approvedAt || '',
        items: o.items.map(function (it) { return { name: it.name, qty: it.qty, unitPrice: it.unitPrice, price: it.price }; })
      };
    }).sort(function (a, b) { return new Date(a.date) - new Date(b.date); });
    return { day: dayStr, orders: list };
  } catch (e) { return { error: e.toString() }; }
}

// Duyệt 1 đơn: ghi người duyệt + thời gian, sau đó MỚI cộng doanh thu vào TongHopThang
function approveOrderRow(startRow, rowCount, approverName, expectDateISO) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    // Ai cũng được duyệt - chỉ cần có tên người duyệt để ghi nhận
    const approverClean = (approverName || '').toString().trim();
    if (!approverClean) return { success: false, error: 'Thiếu tên người duyệt.' };

    const sh = _donHangSheet();
    const row = parseInt(startRow, 10);
    if (!row || row < 2 || row > sh.getLastRow()) return { success: false, error: 'Dòng đơn hàng không hợp lệ.' };
    // [GĐ4] kiểm tra dòng là DÒNG ĐẦU của đúng đơn đó (ranh giới + ngày + SĐT + tổng tiền); web cũ chỉ gửi ngày vẫn chạy
    const fpA = (expectDateISO && typeof expectDateISO === 'object') ? expectDateISO : (expectDateISO ? { d: expectDateISO } : null);
    if (typeof _wgCheckBlock_ === 'function') {
      const chkA = _wgCheckBlock_(sh, row, rowCount, fpA);
      if (!chkA.ok) return { success: false, error: chkA.error };
    }
    const dateVal = sh.getRange(row, 1).getValue();
    if (!dateVal || dateVal === '') return { success: false, error: 'Dòng này không phải dòng đầu của một đơn hàng.' };
    if ((sh.getRange(row, COL_STATUS).getValue() || '').toString() === 'Huy') return { success: false, error: 'Đơn hàng này đã bị hủy/thay thế, không thể duyệt.' };
    const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
    const expISO = (expectDateISO && typeof expectDateISO === 'object') ? expectDateISO.d : expectDateISO;
    if (expISO && Math.abs(d.getTime() - new Date(expISO).getTime()) > 1500) {
      return { success: false, error: 'Đơn hàng đã bị thay đổi - vui lòng tải lại danh sách rồi duyệt lại.' };
    }
    if (!_orderNeedsApproval(d)) return { success: false, error: 'Đơn trước 01/10/2026 không cần duyệt.' };
    const oldRow = sh.getRange(row, COL_APPROVER, 1, 2).getValues()[0];
    const old = (oldRow[0] || '').toString().trim();
    if (old && _approvalValid_(old, oldRow[1])) return { success: false, error: 'Đơn này đã được duyệt bởi ' + old + '.' };

    const n = Math.max(1, parseInt(rowCount, 10) || 1);
    const totals = _readOrderBlockTotals(sh, row, n);
    if (!totals) return { success: false, error: 'Không đọc được dữ liệu đơn hàng.' };

    const now = new Date();
    sh.getRange(row, COL_APPROVER, 1, 2).setValues([[approverClean, now]]);
    try { _tongHopApplyDelta(totals.date, totals.revenue, 1, totals.qty, totals.phone); } catch (e2) { Logger.log('Lỗi cộng TongHopThang khi duyệt: ' + e2); }
    return { success: true, approvedBy: approverClean, approvedAt: now.toISOString() };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally { lock.releaseLock(); }
}