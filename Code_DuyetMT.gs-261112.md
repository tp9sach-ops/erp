/** ============================================================
 *  9 SẠCH - DUYỆT ĐƠN KÊNH MT (ai cũng được duyệt, ghi nhận người duyệt)
 * ============================================================ */
function getMTApprovalList(monthKey) {
  try {
    const parts = (monthKey || '').split('-');
    if (parts.length !== 2) return { error: 'Tháng không hợp lệ' };
    const sh = _mtStatsSheet();
    const data = sh.getDataRange().getValues();
    const out = [];
    data.slice(1).forEach(function (r, i) {
      if (!r[0]) return;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (isNaN(d.getTime()) || _monthKey(d) !== monthKey) return;
      const needs = (r[12] || '').toString() === 'Y';
      const by = (r[10] || '').toString().trim();
      out.push({
        sheetRow: i + 2, date: d.toISOString(), channel: r[1] || '', poNo: (r[2] || '').toString(),
        store: _mtStr_(r[3]), boxes: Number(r[4]) || 0, total: Number(r[7]) || 0, emp: r[9] || '',
        needsApproval: needs, approvedBy: by, approvedAt: _clean(r[11]) || ''
      });
    });
    out.sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
    return { orders: out };
  } catch (e) { return { error: e.toString() }; }
}

function approveMTOrder(sheetRow, channel, poNo, approverName) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    const sh = _mtStatsSheet();
    const row = parseInt(sheetRow, 10);
    if (!row || row < 2 || row > sh.getLastRow()) return { success: false, error: 'Dòng không hợp lệ (có thể đã bị xóa).' };
    const r = sh.getRange(row, 1, 1, 13).getValues()[0];
    // Chống lệch dòng nếu sheet vừa bị xóa/chèn dòng
    if ((r[1] || '') !== channel || (r[2] || '').toString() !== (poNo || '').toString())
      return { success: false, error: 'Dữ liệu đã thay đổi, vui lòng bấm Làm Mới rồi duyệt lại.' };
    if ((r[10] || '').toString().trim()) return { success: false, error: 'Đơn này đã được duyệt bởi ' + r[10] + '.' };
    const name = (approverName || '').toString().trim();
    if (!name) return { success: false, error: 'Thiếu tên người duyệt.' };
    sh.getRange(row, 11, 1, 2).setValues([[name, new Date()]]);
    try { _kpiBust_(); } catch (e0) {}
    return { success: true, approvedBy: name };
  } catch (e) { return { success: false, error: e.toString() };
  } finally { lock.releaseLock(); }
}