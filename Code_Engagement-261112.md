/** ============================================================
 *  9 SẠCH - ĐƠN HÀNG NHÂN VIÊN THEO FANPAGE (Thống Kê Tương Tác)
 *  ================================================================
 *  Nguồn: file Excel "Thống kê tương tác" xuất từ Pancake, gồm 2 sheet:
 *  "all By user" (theo từng nhân viên - dùng tên ảo) và "all By time"
 *  (theo từng ngày, gộp tất cả nhân viên). Client (index.html) đọc file
 *  bằng SheetJS rồi gửi dữ liệu đã parse lên đây để lưu.
 *
 *  CÁCH DÙNG: Trong Apps Script Editor, bấm "+" cạnh Files -> Script ->
 *  đặt tên "Code_Engagement" -> dán TOÀN BỘ nội dung file này vào.
 *  Không cần sửa Code.gs - dùng lại _sheet(), _clean(), _monthKey()
 *  đã có sẵn.
 * ============================================================ */

const ENGAGE_EMP_HEADERS = ['Batch ID','Từ Ngày','Đến Ngày','Tên Ảo','Tên Thật',
  'KH Cũ (TT)','KH Mới (TT)','Tổng TT','Tin Nhắn','Bình Luận','Hội Thoại Mới',
  'Đơn KH Mới','Đơn KH Cũ','Tổng Đơn','Tỉ Lệ CĐ TB (%)','Tỉ Lệ CĐ KH Mới (%)','Tỉ Lệ CĐ KH Cũ (%)',
  'Cập Nhật Lúc','Nhân Viên Upload'];
function _engageEmpSheet(){ return _sheet('Engagement_ByEmployee', ENGAGE_EMP_HEADERS); }

const ENGAGE_DAY_HEADERS = ['Batch ID','Ngày','KH Cũ','KH Mới','Tổng TT','Tin Nhắn','Bình Luận','Hội Thoại Mới',
  'Đơn KH Mới','Đơn KH Cũ','Tổng Đơn','Tỉ Lệ CĐ TB (%)','Tỉ Lệ CĐ KH Mới (%)','Tỉ Lệ CĐ KH Cũ (%)'];
function _engageDaySheet(){ return _sheet('Engagement_Daily', ENGAGE_DAY_HEADERS); }

const ENGAGE_LOG_HEADERS = ['Batch ID','Thời Gian Upload','Tên File','Từ Ngày','Đến Ngày','Số NV','Số Ngày','Nhân Viên Upload'];
function _engageLogSheet(){ return _sheet('Engagement_UploadLog', ENGAGE_LOG_HEADERS); }

// Nhận dữ liệu đã parse sẵn từ client (empRowsArr theo nhân viên, dayRowsArr
// theo ngày gộp toàn hệ thống) và lưu thành 1 lượt upload (batch).
function importEngagementBatch(empRowsArr, dayRowsArr, meta) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    const batchId = 'EG' + new Date().getTime();
    const fromD = meta.fromDate ? new Date(meta.fromDate) : new Date();
    const toD = meta.toDate ? new Date(meta.toDate) : new Date();

    const empSh = _engageEmpSheet();
    const empRows = (empRowsArr || []).map(function (e) {
      return [batchId, fromD, toD, e.virtualName || '', e.realName || '',
      Number(e.khCu) || 0, Number(e.khMoi) || 0, Number(e.tongTT) || 0, Number(e.tinNhan) || 0,
      Number(e.binhLuan) || 0, Number(e.hoiThoaiMoi) || 0, Number(e.donKhMoi) || 0, Number(e.donKhCu) || 0,
      Number(e.tongDon) || 0, Number(e.tiLeCD_TB) || 0, Number(e.tiLeCD_Moi) || 0, Number(e.tiLeCD_Cu) || 0,
        new Date(), meta.employeeName || ''];
    });
    if (empRows.length) {
      const startRow = empSh.getLastRow() + 1;
      empSh.getRange(startRow, 1, empRows.length, ENGAGE_EMP_HEADERS.length).setValues(empRows);
    }

    const daySh = _engageDaySheet();
    const dayRows = (dayRowsArr || []).map(function (d) {
      return [batchId, new Date(d.dateStr), Number(d.khCu) || 0, Number(d.khMoi) || 0, Number(d.tongTT) || 0,
      Number(d.tinNhan) || 0, Number(d.binhLuan) || 0, Number(d.hoiThoaiMoi) || 0, Number(d.donKhMoi) || 0,
      Number(d.donKhCu) || 0, Number(d.tongDon) || 0, Number(d.tiLeCD_TB) || 0, Number(d.tiLeCD_Moi) || 0, Number(d.tiLeCD_Cu) || 0];
    });
    if (dayRows.length) {
      const startRow = daySh.getLastRow() + 1;
      daySh.getRange(startRow, 1, dayRows.length, ENGAGE_DAY_HEADERS.length).setValues(dayRows);
    }

    const logSh = _engageLogSheet();
    logSh.appendRow([batchId, new Date(), meta.fileName || '', fromD, toD, empRows.length, dayRows.length, meta.employeeName || '']);

    return { success: true, batchId: batchId, empCount: empRows.length, dayCount: dayRows.length };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally { lock.releaseLock(); }
}

function getEngagementUploadLog() {
  try {
    const sh = _engageLogSheet();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];
    const data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
    return data.map(function (r) {
      return {
        batchId: (r[0] || '').toString(), time: _clean(r[1]), fileName: r[2] || '',
        from: _clean(r[3]), to: _clean(r[4]), empCount: Number(r[5]) || 0, dayCount: Number(r[6]) || 0, emp: r[7] || ''
      };
    }).sort(function (a, b) { return new Date(b.time) - new Date(a.time); });
  } catch (e) { return []; }
}

// Xóa 1 lượt upload - xóa hết dòng liên quan (nhân viên + theo ngày) + log.
function deleteEngagementBatchServer(batchId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    [_engageEmpSheet(), _engageDaySheet()].forEach(function (sh) {
      const lastRow = sh.getLastRow();
      if (lastRow < 2) return;
      const data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      const rowsToDelete = [];
      data.forEach(function (r, i) { if ((r[0] || '').toString() === batchId) rowsToDelete.push(i + 2); });
      rowsToDelete.sort(function (a, b) { return b - a; });
      rowsToDelete.forEach(function (rowIdx) { sh.deleteRow(rowIdx); });
    });
    const logSh = _engageLogSheet();
    const lastRow = logSh.getLastRow();
    if (lastRow > 1) {
      const data = logSh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < data.length; i++) {
        if ((data[i][0] || '').toString() === batchId) { logSh.deleteRow(i + 2); break; }
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function _engageMonthOverlap(fromD, toD, y, m) {
  const monthStart = new Date(y, m, 1, 0, 0, 0, 0);
  const monthEnd = new Date(y, m + 1, 0, 23, 59, 59, 999);
  return fromD <= monthEnd && toD >= monthStart;
}

function getEngagementAvailableMonths() {
  try {
    const set = {};
    const sh = _engageDaySheet();
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
      data.forEach(function (r) {
        if (!r[1]) return;
        const d = r[1] instanceof Date ? r[1] : new Date(r[1]);
        if (!isNaN(d.getTime())) set[_monthKey(d)] = 1;
      });
    }
    const empSh = _engageEmpSheet();
    const empLastRow = empSh.getLastRow();
    if (empLastRow > 1) {
      const data = empSh.getRange(2, 2, empLastRow - 1, 2).getValues();
      data.forEach(function (r) {
        const from = r[0] instanceof Date ? r[0] : new Date(r[0]);
        const to = r[1] instanceof Date ? r[1] : new Date(r[1]);
        if (isNaN(from.getTime()) || isNaN(to.getTime())) return;
        let cur = new Date(from.getFullYear(), from.getMonth(), 1);
        const end = new Date(to.getFullYear(), to.getMonth(), 1);
        let guard = 0;
        while (cur <= end && guard < 60) { set[_monthKey(cur)] = 1; cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1); guard++; }
      });
    }
    const keys = Object.keys(set).sort().reverse();
    return keys.length ? keys : [_monthKey(new Date())];
  } catch (e) { return [_monthKey(new Date())]; }
}

// Lấy dữ liệu theo nhân viên khớp tháng (nếu 1 nhân viên có nhiều lượt
// upload trùng tháng, lấy lượt cập nhật gần nhất) + tổng hợp toàn hệ
// thống theo NGÀY trong tháng đó (so sánh với tháng trước).
function getEngagementMonthData(monthKey) {
  try {
    const parts = (monthKey || '').split('-');
    if (parts.length !== 2) return { error: 'Tháng không hợp lệ' };
    const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;

    // --- theo nhân viên ---
    const empSh = _engageEmpSheet();
    const empLastRow = empSh.getLastRow();
    const byRealName = {};
    if (empLastRow > 1) {
      const data = empSh.getRange(2, 1, empLastRow - 1, ENGAGE_EMP_HEADERS.length).getValues();
      data.forEach(function (r) {
        const from = r[1] instanceof Date ? r[1] : new Date(r[1]);
        const to = r[2] instanceof Date ? r[2] : new Date(r[2]);
        if (isNaN(from.getTime()) || isNaN(to.getTime())) return;
        if (!_engageMonthOverlap(from, to, y, m)) return;
        const realName = (r[4] || '').toString().trim();
        if (!realName) return;
        const updatedAt = r[17] instanceof Date ? r[17].getTime() : 0;
        if (!byRealName[realName] || updatedAt >= byRealName[realName]._updatedAt) {
          byRealName[realName] = {
            virtualName: (r[3] || '').toString(), realName: realName,
            khCu: Number(r[5]) || 0, khMoi: Number(r[6]) || 0, tongTT: Number(r[7]) || 0,
            tinNhan: Number(r[8]) || 0, binhLuan: Number(r[9]) || 0, hoiThoaiMoi: Number(r[10]) || 0,
            donKhMoi: Number(r[11]) || 0, donKhCu: Number(r[12]) || 0, tongDon: Number(r[13]) || 0,
            tiLeCD_TB: Number(r[14]) || 0, tiLeCD_Moi: Number(r[15]) || 0, tiLeCD_Cu: Number(r[16]) || 0,
            _updatedAt: updatedAt
          };
        }
      });
    }
    const employees = Object.keys(byRealName).map(function (k) { const o = byRealName[k]; delete o._updatedAt; return o; });

    // --- tổng hợp theo ngày (toàn hệ thống) - tháng này & tháng trước ---
    function daySumForMonth(yy, mm) {
      const daySh = _engageDaySheet();
      const lastRow = daySh.getLastRow();
      const sum = { khCu: 0, khMoi: 0, tongTT: 0, tinNhan: 0, binhLuan: 0, hoiThoaiMoi: 0, donKhMoi: 0, donKhCu: 0, tongDon: 0 };
      if (lastRow < 2) return sum;
      const data = daySh.getRange(2, 1, lastRow - 1, ENGAGE_DAY_HEADERS.length).getValues();
      data.forEach(function (r) {
        if (!r[1]) return;
        const d = r[1] instanceof Date ? r[1] : new Date(r[1]);
        if (d.getFullYear() !== yy || d.getMonth() !== mm) return;
        sum.khCu += Number(r[2]) || 0; sum.khMoi += Number(r[3]) || 0; sum.tongTT += Number(r[4]) || 0;
        sum.tinNhan += Number(r[5]) || 0; sum.binhLuan += Number(r[6]) || 0; sum.hoiThoaiMoi += Number(r[7]) || 0;
        sum.donKhMoi += Number(r[8]) || 0; sum.donKhCu += Number(r[9]) || 0; sum.tongDon += Number(r[10]) || 0;
      });
      return sum;
    }
    function pctLocal(cur, prev) { if (!prev) return cur ? 100 : 0; return (cur - prev) / prev * 100; }

    const thisSum = daySumForMonth(y, m);
    let lmY = y, lmM = m - 1; if (lmM < 0) { lmM = 11; lmY--; }
    const lastSum = daySumForMonth(lmY, lmM);

    const tiLeCD_TB = thisSum.tongTT > 0 ? (thisSum.tongDon / thisSum.tongTT * 100) : 0;
    const tiLeCD_TB_last = lastSum.tongTT > 0 ? (lastSum.tongDon / lastSum.tongTT * 100) : 0;

    const summary = {
      khCu: thisSum.khCu, khMoi: thisSum.khMoi, tongTT: thisSum.tongTT,
      tongDon: thisSum.tongDon, tiLeCD_TB: tiLeCD_TB,
      tongTTPct: pctLocal(thisSum.tongTT, lastSum.tongTT),
      khMoiPct: pctLocal(thisSum.khMoi, lastSum.khMoi),
      tongDonPct: pctLocal(thisSum.tongDon, lastSum.tongDon),
      tiLeCD_TBDelta: tiLeCD_TB - tiLeCD_TB_last
    };

    return { monthKey: monthKey, employees: employees, summary: summary };
  } catch (e) { return { error: e.toString() }; }
}

// Tổng số đơn của 1 nhân viên (theo tên thật) trong 1 tháng - có thể
// gọi từ bất kỳ nơi nào khác cần con số này (VD: trang KPI B2B).
function getEngagementOrdersForEmployee(realName, monthKey) {
  try {
    const res = getEngagementMonthData(monthKey);
    if (res.error) return 0;
    const found = (res.employees || []).find(function (e) { return e.realName === realName; });
    return found ? found.tongDon : 0;
  } catch (e) { return 0; }
}

// [MỚI] Lấy dữ liệu tháng đang xem (giống getEngagementMonthData) nhưng
// bổ sung thêm 2 chỉ số TB3T (trung bình 3 tháng liền kề TRƯỚC tháng
// đang xem, không tính tháng hiện tại) cho TỪNG nhân viên:
//   - tb3tTongDon: TB3T của "Tổng Đơn" (đơn hàng đã chốt)
//   - tb3tTyLeCD: TB3T của "Tỷ Lệ Chốt Đơn TB" (%)
// Nếu 1 trong 3 tháng liền kề không có dữ liệu của nhân viên đó thì bỏ
// qua tháng đó khi tính trung bình (không tính là 0). Nếu cả 3 tháng
// đều không có dữ liệu -> trả về null để giao diện hiển thị "Chưa có
// dữ liệu" thay vì hiển thị 0.
function getEngagementMonthDataWithTB3T(monthKey) {
  try {
    const res = getEngagementMonthData(monthKey);
    if (res.error) return res;

    const priorMonthKeys = [1, 2, 3].map(function (i) { return _engagePriorMonthKey(monthKey, i); });
    const priorDataList = priorMonthKeys.map(function (mk) { return getEngagementMonthData(mk); });

    (res.employees || []).forEach(function (emp) {
      const donVals = [];
      const tyLeVals = [];
      priorDataList.forEach(function (pd) {
        if (pd.error) return;
        const found = (pd.employees || []).find(function (e) { return e.realName === emp.realName; });
        if (found) {
          donVals.push(found.tongDon || 0);
          tyLeVals.push(found.tiLeCD_TB || 0);
        }
      });
      emp.tb3tTongDon = donVals.length ? (donVals.reduce(function (a, b) { return a + b; }, 0) / donVals.length) : null;
      emp.tb3tTyLeCD = tyLeVals.length ? (tyLeVals.reduce(function (a, b) { return a + b; }, 0) / tyLeVals.length) : null;
    });

    return res;
  } catch (e) { return { error: e.toString() }; }
}
// ════════════════════════════════════════════════════════════
// [MỚI] TARGET TB 3 THÁNG GẦN NHẤT + DỮ LIỆU 12 THÁNG CHO BIỂU ĐỒ
// ════════════════════════════════════════════════════════════
function _engageSumDaySheetForMonth(y, m) {
  const sh = _engageDaySheet();
  const lastRow = sh.getLastRow();
  const sum = { khCu: 0, khMoi: 0, tongTT: 0, tinNhan: 0, binhLuan: 0, hoiThoaiMoi: 0, donKhMoi: 0, donKhCu: 0, tongDon: 0 };
  if (lastRow < 2) { sum.tiLeCD_TB = 0; return sum; }
  const data = sh.getRange(2, 1, lastRow - 1, ENGAGE_DAY_HEADERS.length).getValues();
  data.forEach(function (r) {
    if (!r[1]) return;
    const d = r[1] instanceof Date ? r[1] : new Date(r[1]);
    if (d.getFullYear() !== y || d.getMonth() !== m) return;
    sum.khCu += Number(r[2]) || 0; sum.khMoi += Number(r[3]) || 0; sum.tongTT += Number(r[4]) || 0;
    sum.tinNhan += Number(r[5]) || 0; sum.binhLuan += Number(r[6]) || 0; sum.hoiThoaiMoi += Number(r[7]) || 0;
    sum.donKhMoi += Number(r[8]) || 0; sum.donKhCu += Number(r[9]) || 0; sum.tongDon += Number(r[10]) || 0;
  });
  sum.tiLeCD_TB = sum.tongTT > 0 ? (sum.tongDon / sum.tongTT * 100) : 0;
  return sum;
}

function _engagePriorMonthKey(monthKey, i) {
  const parts = monthKey.split('-');
  const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;
  let pm = m - i, py = y; while (pm < 0) { pm += 12; py--; }
  return py + '-' + ('0' + (pm + 1)).slice(-2);
}

// Target = trung bình 3 tháng liền kề TRƯỚC tháng đang xem.
function getEngagementTargetAvg3(monthKey) {
  try {
    const sums = [1, 2, 3].map(function (i) {
      const k = _engagePriorMonthKey(monthKey, i);
      const p = k.split('-');
      return _engageSumDaySheetForMonth(parseInt(p[0], 10), parseInt(p[1], 10) - 1);
    });
    function avg(field) { return sums.reduce(function (a, s) { return a + s[field]; }, 0) / 3; }
    return {
      monthKey: monthKey,
      tongTT: avg('tongTT'), khMoi: avg('khMoi'), khCu: avg('khCu'),
      tongDon: avg('tongDon'), tiLeCD_TB: avg('tiLeCD_TB')
    };
  } catch (e) { return { error: e.toString() }; }
}

// Dữ liệu 12 tháng của 1 năm - dùng vẽ biểu đồ xu hướng, client tự chọn chỉ số để hiển thị.
function getEngagementYearlySeries(year) {
  try {
    const y = parseInt(year, 10) || new Date().getFullYear();
    const labels = [], tongTT = [], khMoi = [], khCu = [], tongDon = [], tiLeCD_TB = [];
    for (let m = 0; m < 12; m++) {
      const s = _engageSumDaySheetForMonth(y, m);
      labels.push('Th.' + (m + 1));
      tongTT.push(s.tongTT); khMoi.push(s.khMoi); khCu.push(s.khCu);
      tongDon.push(s.tongDon); tiLeCD_TB.push(Math.round(s.tiLeCD_TB * 100) / 100);
    }
    return { year: y, labels: labels, tongTT: tongTT, khMoi: khMoi, khCu: khCu, tongDon: tongDon, tiLeCD_TB: tiLeCD_TB };
  } catch (e) { return { error: e.toString() }; }
}
// [MỚI] Chuỗi số liệu nhiều tháng của 1 nhân viên (theo tên thật), kèm
// dòng trung bình cộng (TB) của đúng số tháng đang xem - dùng cho modal
// chi tiết nhân viên ở trang Khách TMĐT (xem theo 1/3/6/12 tháng).
function getEngagementEmployeeMonthlySeries(realName, monthsCount, endMonthKey){
  try{
    var n = parseInt(monthsCount,10) || 1;
    if(n<1) n=1; if(n>24) n=24;

    var endY, endM;
    var ek = (endMonthKey||'').toString().trim();
    var mEk = ek.match(/^(\d{4})-(\d{1,2})$/);
    if(mEk){
      endY = parseInt(mEk[1],10);
      endM = parseInt(mEk[2],10)-1;
    } else {
      var now = new Date();
      endY = now.getFullYear();
      endM = now.getMonth();
    }

    var keys = [];
    for(var i=n-1;i>=0;i--){
      var mm=endM-i, yy=endY;
      while(mm<0){ mm+=12; yy--; }
      keys.push(yy+'-'+('0'+(mm+1)).slice(-2));
    }
    var emptyRow = { khCu:0, khMoi:0, tongTT:0, tinNhan:0, binhLuan:0, hoiThoaiMoi:0,
      donKhMoi:0, donKhCu:0, tongDon:0, tiLeCD_TB:0, tiLeCD_Moi:0, tiLeCD_Cu:0 };
    var rows = keys.map(function(k){
      var res = getEngagementMonthData(k);
      var found = (!res.error && res.employees) ? res.employees.find(function(e){ return e.realName===realName; }) : null;
      var base = found ? found : emptyRow;
      var p=k.split('-');
      return {
        monthKey:k, label:'Th.'+p[1]+'/'+p[0],
        khCu:base.khCu||0, khMoi:base.khMoi||0, tongTT:base.tongTT||0, tinNhan:base.tinNhan||0,
        binhLuan:base.binhLuan||0, hoiThoaiMoi:base.hoiThoaiMoi||0, donKhMoi:base.donKhMoi||0,
        donKhCu:base.donKhCu||0, tongDon:base.tongDon||0, tiLeCD_TB:base.tiLeCD_TB||0,
        tiLeCD_Moi:base.tiLeCD_Moi||0, tiLeCD_Cu:base.tiLeCD_Cu||0
      };
    });
    function avg(field){
      if(!rows.length) return 0;
      return rows.reduce(function(a,r){ return a+(r[field]||0); },0)/rows.length;
    }
    var avgRow = {
      monthKey:'TB', label:'TB '+n+' Tháng',
      khCu:avg('khCu'), khMoi:avg('khMoi'), tongTT:avg('tongTT'), tinNhan:avg('tinNhan'),
      binhLuan:avg('binhLuan'), hoiThoaiMoi:avg('hoiThoaiMoi'), donKhMoi:avg('donKhMoi'),
      donKhCu:avg('donKhCu'), tongDon:avg('tongDon'), tiLeCD_TB:avg('tiLeCD_TB'),
      tiLeCD_Moi:avg('tiLeCD_Moi'), tiLeCD_Cu:avg('tiLeCD_Cu')
    };
    return { realName:realName, monthsCount:n, rows:rows, avg:avgRow };
  } catch(e){ return { error:e.toString(), rows:[] }; }
}
// ════════════════════════════════════════════════════════════
// [MỚI] TỔNG HỢP NHIỀU THÁNG CÙNG LÚC (cho phép chọn 2-3 tháng
// hoặc "Chọn Tất Cả" ở trang Khách TMĐT)
// ════════════════════════════════════════════════════════════
function getEngagementMultiMonthData(monthKeysArr) {
  try {
    var keys = (monthKeysArr || []).slice();
    if (!keys.length) keys = [_monthKey(new Date())];
    var empSum = {};
    var teamSum = { khCu:0, khMoi:0, tongTT:0, tongDon:0 };
    keys.forEach(function (k) {
      var res = getEngagementMonthData(k);
      if (res.error) return;
      (res.employees || []).forEach(function (e) {
        if (!empSum[e.realName]) {
          empSum[e.realName] = { virtualName: e.virtualName, realName: e.realName,
            khCu:0, khMoi:0, tongTT:0, tinNhan:0, binhLuan:0, hoiThoaiMoi:0, donKhMoi:0, donKhCu:0, tongDon:0 };
        }
        var acc = empSum[e.realName];
        acc.khCu += e.khCu||0; acc.khMoi += e.khMoi||0; acc.tongTT += e.tongTT||0;
        acc.tinNhan += e.tinNhan||0; acc.binhLuan += e.binhLuan||0; acc.hoiThoaiMoi += e.hoiThoaiMoi||0;
        acc.donKhMoi += e.donKhMoi||0; acc.donKhCu += e.donKhCu||0; acc.tongDon += e.tongDon||0;
      });
      if (res.summary) {
        teamSum.khCu += res.summary.khCu||0; teamSum.khMoi += res.summary.khMoi||0;
        teamSum.tongTT += res.summary.tongTT||0; teamSum.tongDon += res.summary.tongDon||0;
      }
    });
    var employees = Object.keys(empSum).map(function (k) {
      var e = empSum[k];
      e.tiLeCD_TB = e.tongTT>0 ? (e.tongDon/e.tongTT*100) : 0;
      e.tiLeCD_Moi = 0; e.tiLeCD_Cu = 0;
      e.tb3tTongDon = null; e.tb3tTyLeCD = null;
      return e;
    }).sort(function (a, b) { return b.tongDon - a.tongDon; });

    var summary = {
      khCu: teamSum.khCu, khMoi: teamSum.khMoi, tongTT: teamSum.tongTT, tongDon: teamSum.tongDon,
      tiLeCD_TB: teamSum.tongTT>0 ? (teamSum.tongDon/teamSum.tongTT*100) : 0,
      tongTTPct: 0, khMoiPct: 0, tongDonPct: 0, tiLeCD_TBDelta: 0
    };

    return { monthKeys: keys, employees: employees, summary: summary };
  } catch (e) { return { error: e.toString() }; }
}

// ════════════════════════════════════════════════════════════
// [MỚI] CHI TIẾT 1 CHỈ SỐ TEAM-WIDE THEO TỪNG NGÀY TRONG THÁNG +
// 3 THÁNG LIỀN KỀ TRƯỚC + TỔNG + TB3T (bấm vào thẻ KPI để xem)
// ════════════════════════════════════════════════════════════
function getEngagementMetricDailyDetail(monthKey, metricField) {
  try {
    var parts = (monthKey || '').split('-');
    var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;
    var daySh = _engageDaySheet();
    var lastRow = daySh.getLastRow();
    var days = [];
    if (lastRow > 1) {
      var data = daySh.getRange(2, 1, lastRow - 1, ENGAGE_DAY_HEADERS.length).getValues();
      data.forEach(function (r) {
        if (!r[1]) return;
        var d = r[1] instanceof Date ? r[1] : new Date(r[1]);
        if (d.getFullYear() !== y || d.getMonth() !== m) return;
        var khCu = Number(r[2]) || 0, khMoi = Number(r[3]) || 0, tongTT = Number(r[4]) || 0,
            tinNhan = Number(r[5]) || 0, binhLuan = Number(r[6]) || 0, hoiThoaiMoi = Number(r[7]) || 0,
            donKhMoi = Number(r[8]) || 0, donKhCu = Number(r[9]) || 0, tongDon = Number(r[10]) || 0;
        var tiLeCD_TB = tongTT > 0 ? (tongDon / tongTT * 100) : 0;
        days.push({ date: d.toISOString(), khCu:khCu, khMoi:khMoi, tongTT:tongTT, tinNhan:tinNhan,
          binhLuan:binhLuan, hoiThoaiMoi:hoiThoaiMoi, donKhMoi:donKhMoi, donKhCu:donKhCu, tongDon:tongDon, tiLeCD_TB:tiLeCD_TB });
      });
    }
    days.sort(function (a, b) { return new Date(a.date) - new Date(b.date); });

    var monthTotal;
    if (metricField === 'tiLeCD_TB') {
      var sumTT = days.reduce(function (a, r) { return a + r.tongTT; }, 0);
      var sumDon = days.reduce(function (a, r) { return a + r.tongDon; }, 0);
      monthTotal = sumTT > 0 ? (sumDon / sumTT * 100) : 0;
    } else {
      monthTotal = days.reduce(function (a, r) { return a + (r[metricField] || 0); }, 0);
    }

    var priorMonths = [];
    for (var i = 3; i >= 1; i--) {
      var mk = _engagePriorMonthKey(monthKey, i);
      var p = mk.split('-');
      var s = _engageSumDaySheetForMonth(parseInt(p[0], 10), parseInt(p[1], 10) - 1);
      var val = metricField === 'tiLeCD_TB' ? s.tiLeCD_TB : (s[metricField] || 0);
      priorMonths.push({ label: 'Th.' + p[1] + '/' + p[0], value: val });
    }
    var validPrior = priorMonths.filter(function (p) { return p.value > 0; }).map(function (p) { return p.value; });
    var avg3 = validPrior.length ? (validPrior.reduce(function (a, b) { return a + b; }, 0) / validPrior.length) : null;

    return {
      monthKey: monthKey, metricField: metricField,
      days: days.map(function (r) { return { date: r.date, value: metricField === 'tiLeCD_TB' ? r.tiLeCD_TB : (r[metricField] || 0) }; }),
      monthTotal: monthTotal, priorMonths: priorMonths, avg3: avg3
    };
  } catch (e) { return { error: e.toString() }; }
}