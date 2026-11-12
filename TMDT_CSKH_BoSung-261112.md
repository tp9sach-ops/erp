/** ============================================================
 *  9 SẠCH - KHÁCH HÀNG TMĐT + KPI CSKH   v4 (Fix toàn diện)
 *  ================================================================
 *  THAY ĐỔI SO VỚI v3:
 *  - [FIX #5,#6] Target Mục 2 (Tổng đơn chốt) = TB3T tổng đơn TOÀN ĐỘI
 *    CSKH × hệ số, chia đều cho SỐ NHÂN VIÊN hiện có trong danh sách
 *    KPI - không còn tính target riêng theo lịch sử từng cá nhân.
 *  - [FIX #4] "Thực hiện" cá nhân luôn tính TƯƠI từ đơn hàng cache mỗi
 *    lần gọi (không có giá trị tồn từ tháng trước).
 *  - [FIX #7] Trả thêm avgTyLeChotDon3T (tỷ lệ chốt đơn TB 3 tháng) rõ
 *    ràng để client hiển thị thành cột riêng.
 *  - [FIX #1] getKPIBonusListCSKH: mỗi nhân viên xử lý trong try/catch
 *    riêng - 1 người lỗi không làm sập cả bảng.
 *  - [MỚI #8] getCSKHOrderDetail() phục vụ xem chi tiết đơn khi bấm số.
 *  - Phần "KHÁCH HÀNG TMĐT" (sheet KhachHangTMDT, upload/log...) GIỮ
 *    NGUYÊN 100% - không liên quan tới KPI.
 * ============================================================ */

const TMDT_CUST_HEADERS = ['SĐT', 'Họ Tên', 'Địa Chỉ', 'Phường/Xã', 'Quận/Huyện', 'Tỉnh/Thành',
  'Ngày Tạo (Lần Đầu)', 'Lần Mua Cuối', 'Số Lần Mua Hàng', 'Nguồn Khách Hàng', 'Phân Công Cho',
  'Cập Nhật Lúc', 'Batch ID Cuối'];

function _tmdtCustSheet() {
  const sh = _sheet('KhachHangTMDT', TMDT_CUST_HEADERS);
  try { sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).setNumberFormat('@'); } catch (e) {}
  return sh;
}

const TMDT_UPLOAD_LOG_HEADERS = ['Batch ID','Thời Gian Upload','Tên File','Số Khách Thêm Mới','Số Khách Cập Nhật','Nhân Viên Upload'];
function _tmdtUploadLogSheet(){ return _sheet('TMDT_UploadLog', TMDT_UPLOAD_LOG_HEADERS); }

function importTMDTCustomersBatch(rowsArr, fileName, employeeName) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    const sh = _tmdtCustSheet();
    const batchId = 'B' + new Date().getTime();
    const lastRow = sh.getLastRow();
    const existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, TMDT_CUST_HEADERS.length).getValues() : [];
    const phoneIdx = {};
    existing.forEach(function (r, i) {
      const p = (r[0] || '').toString().replace(/\D/g, '');
      if (p) phoneIdx[p] = i + 2;
    });

    let added = 0, updated = 0, skipped = 0;
    const newRows = [];
    (rowsArr || []).forEach(function (c) {
      const phone = (c.phone || '').toString().replace(/\D/g, '');
      if (!phone) { skipped++; return; }
      const createdISO = c.createdDate || '';
      const lastPurISO = c.lastPurchaseDate || createdISO;
      const purchaseCount = Number(c.purchaseCount) || 0;

      if (phoneIdx[phone]) {
        const row = phoneIdx[phone];
        const oldCreated = sh.getRange(row, 7).getValue();
        sh.getRange(row, 2, 1, 5).setValues([[c.name || '', c.address || '', c.ward || '', c.district || '', c.province || '']]);
        if (!oldCreated && createdISO) sh.getRange(row, 7).setValue(new Date(createdISO));
        sh.getRange(row, 8, 1, 4).setValues([[lastPurISO ? new Date(lastPurISO) : '', purchaseCount, c.source || '', c.assignedTo || '']]);
        sh.getRange(row, 12).setValue(new Date());
        sh.getRange(row, 13).setValue(batchId);
        updated++;
      } else {
        newRows.push([
          phone, c.name || '', c.address || '', c.ward || '', c.district || '', c.province || '',
          createdISO ? new Date(createdISO) : new Date(), lastPurISO ? new Date(lastPurISO) : '',
          purchaseCount, c.source || '', c.assignedTo || '', new Date(), batchId
        ]);
        added++;
      }
    });
    if (newRows.length) {
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, newRows.length, 1).setNumberFormat('@');
      sh.getRange(startRow, 1, newRows.length, TMDT_CUST_HEADERS.length).setValues(newRows);
    }

    const logSh = _tmdtUploadLogSheet();
    logSh.appendRow([batchId, new Date(), fileName || '', added, updated, employeeName || '']);

    return { success: true, added: added, updated: updated, skipped: skipped, batchId: batchId };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally { lock.releaseLock(); }
}

function getTMDTUploadLog(){
  try{
    const sh=_tmdtUploadLogSheet();
    const lastRow=sh.getLastRow();
    if(lastRow<2) return [];
    const data=sh.getRange(2,1,lastRow-1,6).getValues();
    return data.map(function(r){
      return { batchId:(r[0]||'').toString(), time:_clean(r[1]), fileName:r[2]||'', added:Number(r[3])||0, updated:Number(r[4])||0, emp:r[5]||'' };
    }).sort(function(a,b){ return new Date(b.time)-new Date(a.time); });
  } catch(e){ return []; }
}

function deleteTMDTUploadBatch(batchId){
  const lock=LockService.getScriptLock();
  try{ lock.waitLock(10000); } catch(e){ return {success:false, error:'Hệ thống đang bận, thử lại sau.'}; }
  try{
    const sh=_tmdtCustSheet();
    const lastRow=sh.getLastRow();
    let deletedCount=0;
    if(lastRow>1){
      const data=sh.getRange(2,1,lastRow-1,TMDT_CUST_HEADERS.length).getValues();
      const rowsToDelete=[];
      data.forEach(function(r,i){
        if((r[12]||'').toString()===batchId) rowsToDelete.push(i+2);
      });
      rowsToDelete.sort(function(a,b){return b-a;});
      rowsToDelete.forEach(function(rowIdx){ sh.deleteRow(rowIdx); deletedCount++; });
    }
    const logSh=_tmdtUploadLogSheet();
    const logLastRow=logSh.getLastRow();
    if(logLastRow>1){
      const logData=logSh.getRange(2,1,logLastRow-1,1).getValues();
      for(let i=0;i<logData.length;i++){
        if((logData[i][0]||'').toString()===batchId){ logSh.deleteRow(i+2); break; }
      }
    }
    return {success:true, deletedCount:deletedCount};
  } catch(e){ return {success:false, error:e.toString()}; }
  finally{ lock.releaseLock(); }
}

function _tmdtLoadAll() {
  const sh = _tmdtCustSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, TMDT_CUST_HEADERS.length).getValues().map(function (r) {
    return {
      phone: (r[0] || '').toString(), name: r[1] || '', address: r[2] || '', ward: r[3] || '',
      district: r[4] || '', province: r[5] || '',
      created: r[6] ? new Date(r[6]) : null, lastPurchase: r[7] ? new Date(r[7]) : null,
      purchaseCount: Number(r[8]) || 0, source: r[9] || '', assignedTo: r[10] || ''
    };
  });
}

function getTMDTCustomerMonthlyTable(year) {
  try {
    const y = parseInt(year, 10) || new Date().getFullYear();
    const all = _tmdtLoadAll();
    const oldCounts = new Array(12).fill(0);
    const newCounts = new Array(12).fill(0);

    all.forEach(function (c) {
      if (c.created && c.created.getFullYear() === y) newCounts[c.created.getMonth()]++;
      if (c.lastPurchase && c.lastPurchase.getFullYear() === y) {
        const sameMonthAsCreated = c.created && c.created.getFullYear() === y && c.created.getMonth() === c.lastPurchase.getMonth();
        if (!sameMonthAsCreated) oldCounts[c.lastPurchase.getMonth()]++;
      }
    });

    function avg3(counts, uptoIdx) {
      const idx = [uptoIdx - 1, uptoIdx - 2, uptoIdx - 3].filter(function (i) { return i >= 0; });
      if (!idx.length) return 0;
      const sum = idx.reduce(function (a, i) { return a + counts[i]; }, 0);
      return sum / idx.length;
    }
    const nowIdx = (y === new Date().getFullYear()) ? new Date().getMonth() : 11;

    return {
      year: y,
      oldRow: { months: oldCounts, avg3: avg3(oldCounts, nowIdx + 1), total: oldCounts.reduce(function (a, b) { return a + b; }, 0) },
      newRow: { months: newCounts, avg3: avg3(newCounts, nowIdx + 1), total: newCounts.reduce(function (a, b) { return a + b; }, 0) }
    };
  } catch (e) { return { error: e.toString() }; }
}

function getTMDTCustomerMonthDetail(year, month, segment) {
  try {
    const y = parseInt(year, 10), m = parseInt(month, 10);
    const all = _tmdtLoadAll();
    const rows = [];
    all.forEach(function (c) {
      if (segment === 'new') {
        if (c.created && c.created.getFullYear() === y && c.created.getMonth() === m) {
          rows.push({ date: c.created.toISOString(), name: c.name,
            address: [c.address, c.ward, c.district, c.province].filter(Boolean).join(', '),
            purchaseCount: c.purchaseCount, phone: c.phone });
        }
      } else {
        if (c.lastPurchase && c.lastPurchase.getFullYear() === y && c.lastPurchase.getMonth() === m) {
          const sameMonthAsCreated = c.created && c.created.getFullYear() === y && c.created.getMonth() === m;
          if (!sameMonthAsCreated) {
            rows.push({ date: c.lastPurchase.toISOString(), name: c.name,
              address: [c.address, c.ward, c.district, c.province].filter(Boolean).join(', '),
              purchaseCount: c.purchaseCount, phone: c.phone });
          }
        }
      }
    });
    rows.sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
    return rows;
  } catch (e) { return []; }
}

// ════════════════════════════════════════════════════════════
//  KPI CSKH - dùng chung sheet KPIThuong (mở rộng thêm cột 17-21)
// ════════════════════════════════════════════════════════════
function _kpiBonusSheetExtended() {
  const sh = _kpiBonusSheet();
  const hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 21)).getValues()[0];
  if (!hdr[16] || hdr[16] !== '(Không dùng - cũ: CSKH DT tháng này)') {
    sh.getRange(1, 17, 1, 5).setValues([[
      '(Không dùng - cũ: CSKH DT tháng này)', '(Không dùng - cũ: CSKH Target DT)', '(Không dùng - cũ: CSKH KH mới TMĐT)',
      '(Không dùng - cũ: CSKH Target KH mới)', 'CSKH Số Lỗi Vi Phạm'
    ]]);
  }
  return sh;
}

// ════════════════════════════════════════════════════════════
//  [FIX] MỤC 2 CSKH: TỔNG ĐƠN HÀNG CHỐT
//  - Target = TB3T tổng đơn TOÀN ĐỘI CSKH × hệ số, chia đều cho SỐ
//    NHÂN VIÊN hiện có trong danh sách KPI.
//  - Thực hiện hiển thị theo TỪNG CÁ NHÂN (minh bạch), tính TƯƠI mỗi
//    lần gọi từ cache đơn hàng của tháng đó - không có giá trị tồn.
// ════════════════════════════════════════════════════════════
function _cskhEmployeeNames(){
  return _getEmployeeListCached().filter(function(e){ return e.role==='CSKH'; })
    .map(function(e){ return e.displayName; });
}

// Tổng đơn TOÀN ĐỘI CSKH trong 1 tháng (dùng để tính Target)
// [CẬP NHẬT] Tổng "Tổng Đơn" TOÀN ĐỘI CSKH trong 1 tháng - lấy từ dữ
// liệu THỐNG KÊ TƯƠNG TÁC (trang "Khách TMĐT", upload từ Pancake),
// KHÔNG dùng đơn hàng tạo trong DonHang nữa. Chỉ cộng những nhân viên
// đang có chức danh CSKH.
function _cskhTeamEngagementTotals(monthKey){
  try{
    const res = getEngagementMonthData(monthKey);
    if(res.error) return { tongTT:0, tongDon:0, khMoi:0, tyLeCD:0 };
    const cskhNames = {};
    _cskhEmployeeNames().forEach(function(n){ cskhNames[(n||'').toString().trim()] = 1; });
    let tongTT=0, tongDon=0, khMoi=0;
    (res.employees||[]).forEach(function(e){
      if(cskhNames[(e.realName||'').toString().trim()]){ tongTT += (e.tongTT||0); tongDon += (e.tongDon||0); khMoi += (e.khMoi||0); }
    });
    // [SỬA] Tỷ lệ chốt đơn = Tổng đơn chốt ÷ Tổng số Lead/KH MỚI nhận
    // (đúng theo văn bản chính sách) - không chia cho Tổng Tương Tác nữa.
    const tyLeCD = khMoi>0 ? (tongDon/khMoi*100) : 0;
    return { tongTT: tongTT, tongDon: tongDon, khMoi: khMoi, tyLeCD: tyLeCD };
  } catch(e){ return { tongTT:0, tongDon:0, khMoi:0, tyLeCD:0 }; }
}

function _cskhTeamOrdersCountForMonth(monthKey){
  return _cskhTeamEngagementTotals(monthKey).tongDon;
}

// Tổng đơn của RIÊNG 1 nhân viên trong 1 tháng - lấy từ dữ liệu THỐNG KÊ
// TƯƠNG TÁC (Khách TMĐT), không phải đơn hàng tạo trong DonHang.
function _cskhIndividualOrdersCountForMonth(monthKey, employeeName){
  try{
    return getEngagementOrdersForEmployee(employeeName, monthKey) || 0;
  } catch(e){ return 0; }
}

// [MỚI] Tỷ lệ chốt đơn CÁ NHÂN của 1 nhân viên trong 1 tháng - lấy từ
// dữ liệu THỐNG KÊ TƯƠNG TÁC (Khách TMĐT).
function _cskhIndividualTyLeChotDon(monthKey, employeeName){
  try{
    const res = getEngagementMonthData(monthKey);
    if(res.error) return 0;
    const found = (res.employees||[]).find(function(e){ return e.realName===employeeName; });
    return found ? (found.tiLeCD_TB||0) : 0;
  } catch(e){ return 0; }
}

// [MỚI] Chi tiết Tổng Đơn Khách TMĐT của TOÀN ĐỘI CSKH trong 1 tháng -
// liệt kê từng nhân viên (dữ liệu nguồn là số tổng hợp từ Pancake, không
// có chi tiết từng đơn lẻ như DonHang).
function getCSKHEngagementTeamDetail(monthKey){
  try{
    const res = getEngagementMonthData(monthKey);
    if(res.error) return [];
    const cskhNames = {};
    _cskhEmployeeNames().forEach(function(n){ cskhNames[(n||'').toString().trim()] = 1; });
    return (res.employees||[])
      .filter(function(e){ return cskhNames[(e.realName||'').toString().trim()]; })
      .map(function(e){ return { realName:e.realName, tongTT:e.tongTT, tongDon:e.tongDon, tiLeCD_TB:e.tiLeCD_TB }; })
      .sort(function(a,b){ return b.tongDon-a.tongDon; });
  } catch(e){ return []; }
}

// [MỚI #8] Chi tiết đơn hàng - nếu truyền employeeName thì chỉ của cá
// nhân đó, không truyền thì lấy toàn đội CSKH.
function getCSKHOrderDetail(monthKey, employeeName){
  try{
    const orders = _getOrdersForMonth(monthKey);
    const target = (employeeName||'').toString().trim();
    let filtered;
    if(target){
      filtered = orders.filter(function(o){ return (o.emp||'').toString().trim()===target; });
    } else {
      const names=_cskhEmployeeNames();
      const set={};
      names.forEach(function(n){set[(n||'').toString().trim()]=1;});
      filtered = orders.filter(function(o){ return set[(o.emp||'').toString().trim()]; });
    }
    return filtered.map(function(o){
      return { date:o.date, name:o.name, phone:o.phone, emp:o.emp, total:o.total, totalNet:_kpiNet_(o.total) };
    }).sort(function(a,b){ return new Date(b.date)-new Date(a.date); });
  } catch(e){ return []; }
}

// Tính Quỹ Mục 2 (Tổng Đơn Chốt) - THEO TOÀN ĐỘI, chia đều đầu người
function _calcCSKHTeamOrderFund(monthKey){
  const nvCount = _getEmployeeListCached().filter(function(e){return e.role==='CSKH';}).length || 1;
  const avgRes = _avgOverAvailableMonths(monthKey, function(mk){ return _cskhTeamEngagementTotals(mk).tongDon; }, 3);
  const heSo = (typeof _kpiHeSoThuong === 'function') ? _kpiHeSoThuong(monthKey) : 1.10;
  const teamTarget = avgRes.hasData ? Math.round(avgRes.avg*heSo) : null;
  const perPersonTarget = teamTarget===null ? null : Math.round(teamTarget/nvCount);

  const teamTotals = _cskhTeamEngagementTotals(monthKey);
  const teamThisRaw = teamTotals.tongDon;
  const isCurMonth = (typeof _isCurrentMonth === 'function') ? _isCurrentMonth(monthKey) : false;
  const teamProjected = isCurMonth ? _projectMetric(teamThisRaw, monthKey) : null;
  const teamForCalc = isCurMonth ? teamProjected : teamThisRaw;

  let tyLeHoanThanh;
  if(teamTarget===null || teamTarget<=0){ tyLeHoanThanh = teamForCalc>0?100:0; }
  else { tyLeHoanThanh = teamForCalc/teamTarget*100; }

  let activatedFund=0, tier='Dưới 80% - Không đạt (0đ)';
  if(tyLeHoanThanh>=100){ activatedFund=5000000; tier='≥100% (100% quỹ)'; }
  else if(tyLeHoanThanh>=90){ activatedFund=4000000; tier='90% - <100% (80% quỹ)'; }
  else if(tyLeHoanThanh>=80){ activatedFund=2500000; tier='80% - <90% (50% quỹ)'; }

  // [CẬP NHẬT] Tỷ lệ chốt đơn lấy từ dữ liệu THỐNG KÊ TƯƠNG TÁC (Khách
  // TMĐT) đã upload, tính riêng trên nhân viên CSKH - không còn dùng
  // ThongKeTrangNgay/DonHang.
  const tyLeChotDonThis = teamTotals.tyLeCD;

  const ratios=[];
  for(let i=1;i<=3;i++){
    const mk=_priorMonthKey(monthKey,i);
    const t = _cskhTeamEngagementTotals(mk);
    if(t.tongTT>0) ratios.push(t.tyLeCD);
  }
  const avgTyLeChotDon3T = ratios.length ? (ratios.reduce(function(a,b){return a+b;},0)/ratios.length) : null;

  let penaltyPoints=0, penaltyAmt=0;
  if(avgTyLeChotDon3T!==null && teamTotals.tongTT>0 && tyLeChotDonThis < avgTyLeChotDon3T){
    penaltyPoints = avgTyLeChotDon3T - tyLeChotDonThis;
    penaltyAmt = Math.round(penaltyPoints*500000);
  }
  const finalTeamFund = Math.max(0, activatedFund - penaltyAmt);
  const perPersonFund = Math.round(finalTeamFund / nvCount);

  return {
    nvCount: nvCount, teamTarget: teamTarget, perPersonTarget: perPersonTarget,
    teamThisRaw: teamThisRaw, teamProjected: teamProjected, isCurrentMonth: isCurMonth,
    tyLeHoanThanh: tyLeHoanThanh, tier: tier, activatedFund: activatedFund,
    tyLeChotDonThis: tyLeChotDonThis, avgTyLeChotDon3T: avgTyLeChotDon3T,
    penaltyPoints: penaltyPoints, penaltyAmt: penaltyAmt,
    finalTeamFund: finalTeamFund, perPersonFund: perPersonFund,
    hasEnoughData: avgRes.hasData, monthsUsed: avgRes.monthsUsed,
    leadDataAvailable: teamTotals.tongTT>0
  };
}

// Giữ nguyên tên hàm public cũ để tương thích - trả về dữ liệu quỹ đội.
function getCSKHOrderFundInfo(monthKey){
  try{ return _calcCSKHTeamOrderFund(monthKey); } catch(e){ return {error:e.toString()}; }
}

// ════════════════════════════════════════════════════════════
//  [CHÍNH SÁCH CSKH MỚI]
//  Mục 1: Quỹ theo Target chung kênh sỉ (5.000.000đ) chia theo HỆ SỐ CA
//         (ca sáng = 6, ca tối = 4): thưởng cá nhân = quỹ kích hoạt × hệ số ca ÷ tổng hệ số cả đội.
//  Mục 2: Tốc độ phản hồi (tối đa 300.000đ) lấy từ file thống kê nhân viên Fanpage/Zalo:
//         <3 phút = 300.000đ | 3 - <5 phút = 150.000đ | từ 5 phút = 0đ.
//  Thưởng thực nhận = (Mục 1 + Mục 2 - Khấu trừ vi phạm) × % Ngày công
// ════════════════════════════════════════════════════════════
const CSKH_HESO_CA = { Sang: 6, Toi: 4 };
function _cskhHeSoCa(ca){ return ca==='Toi' ? CSKH_HESO_CA.Toi : CSKH_HESO_CA.Sang; }

function _cskhSpeedBonus(avgSeconds){
  if(avgSeconds===null || avgSeconds===undefined || isNaN(avgSeconds)) return { bonus:0, tier:'Chưa có dữ liệu' };
  if(avgSeconds < 180) return { bonus:300000, tier:'Dưới 3 phút (100%)' };
  if(avgSeconds < 300) return { bonus:150000, tier:'Từ 3 đến dưới 5 phút (50%)' };
  return { bonus:0, tier:'Từ 5 phút trở lên (0%)' };
}

function _calcCSKHBonus(input) {
  const salesBonus = Math.round(Number(input.kenhSiShare)||0);
  const sp = _cskhSpeedBonus(input.avgResponseSeconds);
  const speedBonus = sp.bonus;

  const violations = Number(input.cskhViolations) || 0;
  const personalScore = Math.max(0, 100 - 10*violations);
  const deduction = violations * 100000;
  const rawTotal = salesBonus + speedBonus;
  let finalTotal, blocked = false;
  if (personalScore < 60) { finalTotal = 0; blocked = true; }
  else { finalTotal = Math.max(0, rawTotal - deduction); }

  const heSoNgayCong = _heSoNgayCong(input.gioCongThucTe);
  finalTotal = Math.round(finalTotal * heSoNgayCong);

  return { salesBonus: salesBonus, speedBonus: speedBonus, speedTier: sp.tier,
    ca: input.ca==='Toi'?'Toi':'Sang', heSoCa: _cskhHeSoCa(input.ca),
    violations: violations, personalScore: personalScore, blocked: blocked,
    deduction: deduction, rawTotal: rawTotal, finalTotal: finalTotal,
    heSoNgayCong: heSoNgayCong };
}

// Đọc ca làm việc đã lưu của toàn bộ CSKH trong tháng -> { tên: 'Sang'|'Toi' }
function _cskhSavedMap(monthKey){
  const sh = _kpiBonusSheetExtended();
  _kpiBonusSheetGioCong();
  const lastRow = sh.getLastRow();
  const savedRows = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, KPI_COL_CA).getValues() : [];
  const savedMap = {};
  savedRows.forEach(function (r) {
    if (_kpiMonthKeyOf(r[0]) === monthKey) {
      savedMap[(r[1] || '').toString()] = {
        note: r[14] || '', cskhViolations: Number(r[20]) || 0,
        gioCongThucTe: Number(r[KPI_COL_GIOCONG-1]) || 208,
        ca: (r[KPI_COL_CA-1] || '').toString()==='Toi' ? 'Toi' : 'Sang'
      };
    }
  });
  return savedMap;
}
// Tổng hệ số của cả đội CSKH + giá trị 1 hệ số
function _cskhShiftSummary(monthKey, employees, savedMap, quyCSKH){
  let sang=0, toi=0;
  employees.forEach(function(e){
    const s = savedMap[e.displayName];
    if(s && s.ca==='Toi') toi++; else sang++;
  });
  const totalHeSo = sang*CSKH_HESO_CA.Sang + toi*CSKH_HESO_CA.Toi;
  const valuePerHeSo = totalHeSo>0 ? (quyCSKH/totalHeSo) : 0;
  return { sang:sang, toi:toi, totalHeSo:totalHeSo, valuePerHeSo:valuePerHeSo };
}

function calcCSKHBonusPreview(input) {
  try {
    const monthKey = input.monthKey || _monthKey(new Date());
    const empName = input.empName || '';
    const employees = _getEmployeeListCached().filter(function(e){return e.role==='CSKH';});
    const savedMap = _cskhSavedMap(monthKey);
    // Ca đang chọn trong form được ưu tiên cho chính nhân viên này
    const mapForCalc = JSON.parse(JSON.stringify(savedMap));
    if(empName) mapForCalc[empName] = Object.assign({}, mapForCalc[empName]||{}, { ca: input.ca==='Toi'?'Toi':'Sang' });
    const kenhSiFund = _calcKenhSiFund(monthKey);
    const shift = _cskhShiftSummary(monthKey, employees, mapForCalc, kenhSiFund.quyCSKH);
    const sp = getResponseSpeedForEmployee(monthKey, empName);
    const calc = _calcCSKHBonus({
      kenhSiShare: shift.valuePerHeSo * _cskhHeSoCa(input.ca),
      avgResponseSeconds: sp ? sp.avgSeconds : null,
      cskhViolations: input.cskhViolations, gioCongThucTe: input.gioCongThucTe, ca: input.ca
    });
    calc.kenhSiFund = kenhSiFund; calc.shift = shift; calc.speed = sp;
    return calc;
  } catch (e) { return { error: e.toString() }; }
}

function saveCSKHBonusInput(monthKey, empName, input) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    const sh = _kpiBonusSheetExtended();
    _kpiBonusSheetGioCong(); // đảm bảo cột 22 + 23 đã có header
    const role = (_getEmployeeListCached().find(function (e) { return e.displayName === empName; }) || {}).role || 'CSKH';
    const baseRow = [monthKey, empName, role, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, input.note || '', new Date()];
    const extRow = [0, 0, 0, 0, Number(input.cskhViolations) || 0];
    const gioCong = Number(input.gioCongThucTe) || 208;
    const ca = input.ca==='Toi' ? 'Toi' : 'Sang';
    let idx = -1;
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        if (_kpiMonthKeyOf(data[i][0]) === monthKey && (data[i][1] || '').toString() === empName) { idx = i + 2; break; }
      }
    }
    let targetRow;
    if (idx > 0) {
      sh.getRange(idx, 1).setNumberFormat('@');
      sh.getRange(idx, 1, 1, baseRow.length).setValues([baseRow]);
      sh.getRange(idx, 17, 1, extRow.length).setValues([extRow]);
      targetRow = idx;
    } else {
      const newRow = sh.getLastRow() + 1;
      sh.getRange(newRow, 1).setNumberFormat('@');
      sh.getRange(newRow, 1, 1, baseRow.length).setValues([baseRow]);
      sh.getRange(newRow, 17, 1, extRow.length).setValues([extRow]);
      targetRow = newRow;
    }
    sh.getRange(targetRow, KPI_COL_GIOCONG).setValue(gioCong);
    sh.getRange(targetRow, KPI_COL_CA).setValue(ca);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getKPIBonusListCSKH(monthKey) {
  try {
    _KPI_ORDERS_CACHE = {}; _KPI_EMP_LIST_CACHE = null;
    const employees = _getEmployeeListCached().filter(function (e) { return e.role === 'CSKH'; });
    if (!employees.length) return [];

    const savedMap = _cskhSavedMap(monthKey);
    const kenhSiFund = _calcKenhSiFund(monthKey);
    const shift = _cskhShiftSummary(monthKey, employees, savedMap, kenhSiFund.quyCSKH);
    const speedMap = getResponseSpeedMap(monthKey);

    const results = [];
    employees.forEach(function (e) {
      try {
        const saved = savedMap[e.displayName];
        const ca = saved ? saved.ca : 'Sang';
        const heSoCa = _cskhHeSoCa(ca);
        const sp = speedMap[(e.displayName||'').toString().trim()] || null;
        const calcInput = {
          kenhSiShare: shift.valuePerHeSo * heSoCa,
          avgResponseSeconds: sp ? sp.avgSeconds : null,
          cskhViolations: saved ? saved.cskhViolations : 0,
          gioCongThucTe: saved ? saved.gioCongThucTe : 208, ca: ca
        };
        const calc = _calcCSKHBonus(calcInput);
        results.push({
          name: e.displayName, role: e.role,
          input: { cskhViolations: calcInput.cskhViolations, note: saved ? saved.note : '', gioCongThucTe: calcInput.gioCongThucTe, ca: ca },
          calc: calc, kenhSiFund: kenhSiFund, shift: shift, speed: sp, ca: ca, heSoCa: heSoCa, error: null
        });
      } catch (errEmp) {
        results.push({ name: e.displayName, role: e.role, error: errEmp.toString(),
          input:{cskhViolations:0,note:'',gioCongThucTe:208,ca:'Sang'},
          calc:{salesBonus:0,speedBonus:0,speedTier:'',ca:'Sang',heSoCa:6,violations:0,personalScore:100,blocked:false,deduction:0,rawTotal:0,finalTotal:0},
          kenhSiFund:kenhSiFund, shift:shift, speed:null, ca:'Sang', heSoCa:6 });
      }
    });
    return results.sort(function (a, b) { return (b.calc.finalTotal||0) - (a.calc.finalTotal||0); });
  } catch (e) { return { error: e.toString() }; }
}