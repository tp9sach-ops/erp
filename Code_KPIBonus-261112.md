/** ============================================================
 *  9 SẠCH - THƯỞNG KPI B2B / QUẢN LÝ B2B  (v6 - Fix toàn diện)
 *  ================================================================
 *  THAY ĐỔI SO VỚI v5:
 *  - [FIX #12] _kenhSiRevenueOfMonth: đổi thành TỔNG TOÀN BỘ doanh thu
 *    DonHang của tháng đó (đúng nghĩa "toàn hệ thống" theo Chính Sách
 *    V13) - KHÔNG còn lọc theo tên nhân viên hiện tại (gây thiếu doanh
 *    thu khi nhân viên đổi vai trò/nghỉ việc).
 *  - [FIX #9] Thêm cache đơn hàng theo tháng (_getOrdersForMonth) dùng
 *    chung cho toàn bộ tính toán trong 1 lần gọi - giảm số lần quét
 *    sheet DonHang từ "N nhân viên x nhiều tháng" xuống còn đúng số
 *    tháng cần (tối đa 4 lần/lượt tải trang).
 *  - Thêm getKenhSiRevenueDetail() + getCSKHOrderDetail() phục vụ
 *    tính năng "xem chi tiết khi bấm vào số" (#8).
 * ============================================================ */

const KPI_BONUS_HEADERS = ['Tháng','Nhân viên','Chức danh','Mục tiêu KH mới','Số KH mới đạt được',
  '(Không dùng - cũ: DT hoàn thành)','(Không dùng - cũ: Target DT TB3T)','Target Dự Án MT (TB3T×1.2)','Dự Án MT Hoàn Thành',
  'Vi phạm 1','Vi phạm 2','Vi phạm 3','Vi phạm 4','KPI đội ngũ (%)','Ghi chú','Cập nhật lúc'];

function _kpiBonusSheet(){
  const sh = _sheet('KPIThuong', KPI_BONUS_HEADERS);
  try {
    const lastRow = sh.getLastRow();
    const numRows = Math.max(lastRow - 1, 1);
    const _c = CacheService.getScriptCache(); if (!_c.get('KPI_FMT')) { sh.getRange(2, 1, numRows, 1).setNumberFormat('@'); _c.put('KPI_FMT', '1', 21600); }
  } catch (e) {}
  return sh;
}

// [MỚI] Cột 22 dùng chung B2B/QLB2B/CSKH - lưu số giờ công thực tế
// trong tháng của từng nhân viên (mặc định 208 giờ = 26 công chuẩn).
const KPI_COL_GIOCONG = 22;
const KPI_COL_CA = 23; // [MỚI] Ca làm việc của CSKH: 'Sang' (hệ số 6) / 'Toi' (hệ số 4)
function _kpiBonusSheetGioCong(){
  const sh = _kpiBonusSheet();
  const hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), KPI_COL_CA)).getValues()[0];
  if (!hdr[KPI_COL_GIOCONG-1] || hdr[KPI_COL_GIOCONG-1] !== 'Giờ Công Thực Tế') {
    sh.getRange(1, KPI_COL_GIOCONG).setValue('Giờ Công Thực Tế');
  }
  if (!hdr[KPI_COL_CA-1] || hdr[KPI_COL_CA-1] !== 'Ca Làm Việc (CSKH)') {
    sh.getRange(1, KPI_COL_CA).setValue('Ca Làm Việc (CSKH)');
  }
  return sh;
}
// [MỚI] Điểm hành vi trung bình (thang 100, sau khi trừ lỗi) của TOÀN BỘ NVKD B2B + CSKH
// trong tháng - dùng cho Mục 4 "Quản trị & tiêu chuẩn vận hành" của Quản lý B2B.
function _teamBehaviorAvgPct(monthKey){
  const emps = _getEmployeeListCached().filter(function(e){ return e.role==='B2B' || e.role==='CSKH'; });
  if(!emps.length) return 100;
  const sh = _kpiBonusSheetGioCong();
  const lastRow = sh.getLastRow();
  const rows = lastRow>1 ? sh.getRange(2,1,lastRow-1,21).getValues() : [];
  const map = {};
  rows.forEach(function(r){
    if(_kpiMonthKeyOf(r[0])===monthKey){
      map[(r[1]||'').toString()] = {
        b2b: (Number(r[9])||0)+(Number(r[10])||0)+(Number(r[11])||0)+(Number(r[12])||0),
        cskh: Number(r[20])||0
      };
    }
  });
  let sum = 0;
  emps.forEach(function(e){
    const m = map[e.displayName];
    const v = m ? (e.role==='B2B' ? m.b2b : m.cskh) : 0;
    sum += Math.max(0, 100 - 10*v);
  });
  return sum/emps.length;
}
function _heSoNgayCong(gioCongThucTe){
  return Math.min(1, (Number(gioCongThucTe)||208)/208);
}

function _kpiMonthKeyOf(val){
  if (val instanceof Date) return _monthKey(val);
  return (val || '').toString().trim();
}

function _kpiBonusFindRow(sh, monthKey, empName){
  const lastRow=sh.getLastRow();
  if(lastRow<2) return -1;
  const data=sh.getRange(2,1,lastRow-1,2).getValues();
  for(let i=0;i<data.length;i++){
    if(_kpiMonthKeyOf(data[i][0])===monthKey && (data[i][1]||'').toString()===empName) return i+2;
  }
  return -1;
}

function saveKPIBonusInput(monthKey, empName, input){
  const lock=LockService.getScriptLock();
  try{ lock.waitLock(8000); } catch(e){ return {success:false, error:'Hệ thống đang bận, thử lại sau.'}; }
  try{
    const sh=_kpiBonusSheetGioCong();
    const role = (_getEmployeeListCached().find(function(e){return e.displayName===empName;})||{}).role || '';
    const row=[
      monthKey, empName, role,
      Number(input.targetNewCust)||0, Number(input.actualNewCust)||0,
      0, 0,
      Number(input.projectTarget)||0, Number(input.projectCompleted)||0,
      Number(input.viol1)||0, Number(input.viol2)||0, Number(input.viol3)||0, Number(input.viol4)||0,
      Number(input.teamPct)||0, input.note||'', new Date()
    ];
    const gioCong = Number(input.gioCongThucTe)||208;
    let targetRow;
    const idx=_kpiBonusFindRow(sh, monthKey, empName);
    if(idx>0){
      sh.getRange(idx,1).setNumberFormat('@');
      sh.getRange(idx,1,1,row.length).setValues([row]);
      targetRow = idx;
    } else {
      const newRow = sh.getLastRow()+1;
      sh.getRange(newRow,1).setNumberFormat('@');
      sh.getRange(newRow,1,1,row.length).setValues([row]);
      targetRow = newRow;
    }
    sh.getRange(targetRow, KPI_COL_GIOCONG).setValue(gioCong);
    return {success:true};
  } catch(e){ return {success:false, error:e.toString()}; }
  finally{ lock.releaseLock(); }
}

// ════════════════════════════════════════════════════════════
//  [MỚI #9] CACHE TRONG 1 LẦN THỰC THI - giảm số lần quét sheet
// ════════════════════════════════════════════════════════════
var _KPI_ORDERS_CACHE = {};
function _getOrdersForMonth(monthKey){
  if (_KPI_ORDERS_CACHE.hasOwnProperty(monthKey)) return _KPI_ORDERS_CACHE[monthKey];
  const parts=monthKey.split('-');
  const y=parseInt(parts[0],10), m=parseInt(parts[1],10)-1;
  const from=new Date(y,m,1,0,0,0,0), to=new Date(y,m+1,0,23,59,59,999);
  const orders=_readDonHangGrouped(from.toISOString(), to.toISOString(), null, true); // chỉ đơn đã được tính (từ 01/10/2026 phải được duyệt)
  _KPI_ORDERS_CACHE[monthKey]=orders;
  return orders;
}

var _KPI_MT_TOTAL_CACHE = {};
function _mtTotalRevenueOfMonth(monthKey){
  if (_KPI_MT_TOTAL_CACHE.hasOwnProperty(monthKey)) return _KPI_MT_TOTAL_CACHE[monthKey];
  const _ckMT = _kpiPastKey_('MT', monthKey);
  if (_ckMT) { const _hit = _kpiCacheNumGet_(_ckMT); if (_hit !== null) { _KPI_MT_TOTAL_CACHE[monthKey] = _hit; return _hit; } }
  try{
    const parts=monthKey.split('-');
    const y=parseInt(parts[0],10), m=parseInt(parts[1],10)-1;
    const from=new Date(y,m,1,0,0,0,0), to=new Date(y,m+1,0,23,59,59,999);
    const rows=getMTStats(from.toISOString(), to.toISOString());
    const total = rows.reduce(function(a,r){ return a+(r.total||0); },0);
    _KPI_MT_TOTAL_CACHE[monthKey]=total;
    if (_ckMT) _kpiCacheNumPut_(_ckMT, total);
    return total;
  } catch(e){ return 0; }
}

// [VAT T11/2026] Doanh thu MT CHƯA VAT của 1 tháng (xem KPI_VAT.gs). Tháng cũ vẫn dùng _mtTotalRevenueOfMonth (có VAT).
var _KPI_MT_NET_CACHE = {};
function _mtNetRevenueOfMonth(monthKey){
  if (_KPI_MT_NET_CACHE.hasOwnProperty(monthKey)) return _KPI_MT_NET_CACHE[monthKey];
  const _ck = _kpiPastKey_('MTN', monthKey);
  if (_ck) { const _hit = _kpiCacheNumGet_(_ck); if (_hit !== null) { _KPI_MT_NET_CACHE[monthKey] = _hit; return _hit; } }
  try{
    const parts=monthKey.split('-');
    const y=parseInt(parts[0],10), m=parseInt(parts[1],10)-1;
    const from=new Date(y,m,1,0,0,0,0), to=new Date(y,m+1,0,23,59,59,999);
    const rows=getMTStats(from.toISOString(), to.toISOString());
    const total = Math.round(rows.reduce(function(a,r){ return a+_kpiMtRowNet_(r); },0));
    _KPI_MT_NET_CACHE[monthKey]=total;
    if (_ck) _kpiCacheNumPut_(_ck, total);
    return total;
  } catch(e){ return 0; }
}

var _KPI_EMP_LIST_CACHE = null;
function _getEmployeeListCached(){
  if(_KPI_EMP_LIST_CACHE) return _KPI_EMP_LIST_CACHE;
  _KPI_EMP_LIST_CACHE = getEmployeeList();
  return _KPI_EMP_LIST_CACHE;
}

// ════════════════════════════════════════════════════════════
//  TRUNG BÌNH CHỈ TÍNH TRÊN CÁC THÁNG CÓ DỮ LIỆU (>0)
// ════════════════════════════════════════════════════════════
function _avgOverAvailableMonths(monthKey, valueFn, monthsBack){
  monthsBack = monthsBack || 3;
  const vals=[];
  for(let i=1;i<=monthsBack;i++){
    const v = valueFn(_priorMonthKey(monthKey, i));
    if(v>0) vals.push(v);
  }
  if(!vals.length) return { avg:null, monthsUsed:0, hasData:false };
  const sum = vals.reduce(function(a,b){return a+b;},0);
  return { avg: sum/vals.length, monthsUsed: vals.length, hasData:true };
}

function _avgOverAvailableMonthsDetail(monthKey, valueFn, monthsBack){
  monthsBack = monthsBack || 3;
  const months = [];
  for(let i=monthsBack;i>=1;i--){
    const mk = _priorMonthKey(monthKey, i);
    months.push({ monthKey: mk, value: valueFn(mk) });
  }
  const validVals = months.filter(function(m){ return m.value>0; }).map(function(m){ return m.value; });
  const hasData = validVals.length>0;
  const avg = hasData ? (validVals.reduce(function(a,b){return a+b;},0)/validVals.length) : null;
  return { months: months, avg: avg, monthsUsed: validVals.length, hasData: hasData };
}

// ════════════════════════════════════════════════════════════
//  [GIỮ LẠI - không dùng nữa cho Kênh Sỉ nhưng có thể tái sử dụng]
// ════════════════════════════════════════════════════════════
function _empRevenueOfMonthKPI(monthKey){
  const orders = _getOrdersForMonth(monthKey);
  const byEmp={};
  orders.forEach(function(o){ if(!o.emp||!o.emp.trim()) return; byEmp[o.emp]=(byEmp[o.emp]||0)+o.total; });
  return byEmp;
}
function _groupRevenueTotalForMonth(monthKey, employeeNames){
  const map=_empRevenueOfMonthKPI(monthKey);
  let sum=0;
  employeeNames.forEach(function(n){ sum+=map[n]||0; });
  return sum;
}
function _kenhSiEmployeeNames(){
  const all = _getEmployeeListCached();
  return all.filter(function(e){ return e.role==='B2B' || e.role==='QLB2B' || e.role==='CSKH'; })
            .map(function(e){ return e.displayName; });
}

// [CẬP NHẬT CHÍNH SÁCH MỚI] Một khách chỉ được tính là KH MỚI khi đạt MỘT trong hai điều kiện:
//   - Đơn đầu tiên có giá trị TRÊN 3.000.000đ
//   - Hoặc khách đã đặt đến đơn thứ 3
// (vẫn tôn trọng phân loại thủ công Cũ/Mới ở DataKhachSi cột 19)
const NEW_CUST_MIN_FIRST_ORDER = 3000000;
const NEW_CUST_MIN_ORDERS = 3;
function _newCustomersByEmployeeForMonth(monthKey){
  const orders=_getOrdersForMonth(monthKey).slice().sort(function(a,b){ return new Date(a.date)-new Date(b.date); });
  const parts=monthKey.split('-');
  const y=parseInt(parts[0],10), m=parseInt(parts[1],10)-1;
  const from=new Date(y,m,1,0,0,0,0), to=new Date(y,m+1,0,23,59,59,999);
  let firstOrderMap=_kpiFirstOrderMapCached_();
  const manualMap=_loadCustomerManualSegMap();
  const info={};
  orders.forEach(function(o){
    const phone=(o.phone||'').toString().replace(/\D/g,'');
    if(!phone) return;
    const emp=o.emp||'(Không rõ)';
    const key=emp+'|'+phone;
    if(!info[key]) info[key]={ emp:emp, phone:phone, n:0, firstTotal:o.total };
    info[key].n++;
  });
  const byEmp={};
  Object.keys(info).forEach(function(key){
    const it=info[key];
    const manual = manualMap[it.phone] || '';
    let isNew;
    if(manual==='Cu') isNew=false;
    else if(manual==='Moi') isNew=true;
    else {
      if (firstOrderMap[it.phone] === undefined && !_KPI_FO_FRESH) { firstOrderMap = _kpiFirstOrderMapCached_(true); }
      const first=firstOrderMap[it.phone];
      isNew = first ? (first>=from && first<=to) : false;
    }
    const qualified = it.firstTotal > NEW_CUST_MIN_FIRST_ORDER || it.n >= NEW_CUST_MIN_ORDERS;
    if(isNew && qualified) byEmp[it.emp]=(byEmp[it.emp]||0)+1;
  });
  return byEmp;
}

// [MỚI - mục 6] Khi xuất văn bản, người dùng có thể chọn "Chỉ tính Lũy Kế
// Thực Tế" - cờ này bật lên thì _projectMetric() bỏ qua phần dự kiến,
// luôn trả về đúng số thực tế đã có tới thời điểm xuất, dù đang là tháng
// hiện tại. Chỉ nên bật/tắt trong đúng 1 lần gọi xuất văn bản (set rồi
// reset ngay trong try/finally ở hàm xuất).
var _KPI_FORCE_ACTUAL_ONLY = false;

function _priorMonthKey(monthKey, i){
  const parts=monthKey.split('-');
  const y=parseInt(parts[0],10), m=parseInt(parts[1],10)-1;
  let pm=m-i, py=y; while(pm<0){pm+=12;py--;}
  return py+'-'+('0'+(pm+1)).slice(-2);
}

function _isCurrentMonth(monthKey){
  return monthKey === _monthKey(new Date());
}

function _projectMetric(actualSoFar, monthKey){
  if(_KPI_FORCE_ACTUAL_ONLY) return actualSoFar;
  if(!_isCurrentMonth(monthKey)) return actualSoFar;
  const now = new Date();
  const parts = monthKey.split('-');
  const y = parseInt(parts[0],10), m = parseInt(parts[1],10)-1;
  const daysInMonth = new Date(y, m+1, 0).getDate();
  const daysElapsed = now.getDate();
  if(!daysElapsed) return actualSoFar;
  return (actualSoFar/daysElapsed) * daysInMonth;
}

function _kpiHeSoThuong(monthKey){
  try{
    if (typeof getThangCaoDiem === 'function') {
      const cfg = getThangCaoDiem(monthKey);
      if (cfg && cfg.pctTang) return 1 + (Number(cfg.pctTang)||0)/100;
    }
  } catch(e){}
  return 1.10;
}
function _kpiHeSoTangTruongMT(monthKey){
  try{
    if (typeof getThangCaoDiem === 'function') {
      const cfg = getThangCaoDiem(monthKey);
      if (cfg && cfg.pctTang) return 1 + (Number(cfg.pctTang)||0)/100;
    }
  } catch(e){}
  return 1.20;
}

function getMTGrowthForMonth(monthKey){
  try{
    const parts=(monthKey||'').split('-');
    if(parts.length!==2) return { mtThis:0, mtTargetAvg3:0, completionPct:0, error:'Tháng không hợp lệ' };
    // [VAT T11/2026] từ T11/2026: Target = TB3T (chưa VAT) × hệ số; doanh thu tính % cũng chưa VAT. Tháng cũ: như cũ.
    const vat = _kpiVatOn_(monthKey);
    const valFn = vat ? _mtNetRevenueOfMonth : _mtTotalRevenueOfMonth;
    const avgRes = _avgOverAvailableMonths(monthKey, valFn, 3);
    const heSo = _kpiHeSoTangTruongMT(monthKey);
    const mtTargetAvg3 = avgRes.hasData ? (avgRes.avg*heSo) : null;
    const mtThis=valFn(monthKey);
    const mtThisGross = vat ? _mtTotalRevenueOfMonth(monthKey) : mtThis;
    const isCurMonth = _isCurrentMonth(monthKey);
    const mtProjected = isCurMonth ? _projectMetric(mtThis, monthKey) : null;
    const mtProjectedGross = isCurMonth ? (vat ? _projectMetric(mtThisGross, monthKey) : mtProjected) : null;
    const mtForCalc = isCurMonth ? mtProjected : mtThis;
    let completionPct;
    if(mtTargetAvg3===null || mtTargetAvg3<=0){ completionPct = mtForCalc>0?100:0; }
    else { completionPct = mtForCalc/mtTargetAvg3*100; }
    return {
      // mtThis / mtProjected / mtTargetAvg3 = số DÙNG ĐỂ TÍNH (từ T11/2026 là CHƯA VAT)
      mtThis:mtThis, mtProjected: mtProjected, isCurrentMonth: isCurMonth,
      mtTargetAvg3: mtTargetAvg3===null?0:mtTargetAvg3, completionPct:completionPct,
      hasEnoughData: avgRes.hasData, monthsUsed: avgRes.monthsUsed,
      vatApplied: vat, vatRate: KPI_VAT_RATE,
      mtThisGross: mtThisGross, mtThisNet: vat ? mtThis : null,
      mtProjectedGross: mtProjectedGross, mtProjectedNet: vat ? mtProjected : null
    };
  } catch(e){ return { mtThis:0, mtTargetAvg3:0, completionPct:0, error:e.toString() }; }
}

// ════════════════════════════════════════════════════════════
//  [FIX #12] DOANH THU KÊNH SỈ = TOÀN BỘ DOANH THU DONHANG THÁNG ĐÓ
//  (đúng nghĩa "toàn hệ thống" - KHÔNG lọc theo tên nhân viên, tránh
//  thiếu doanh thu khi nhân viên đổi vai trò/nghỉ việc)
// ════════════════════════════════════════════════════════════
function _kenhSiRevenueOfMonth(monthKey){
  const ck = _kpiPastKey_('KS', monthKey);
  if (ck) { const hit = _kpiCacheNumGet_(ck); if (hit !== null) return hit; }
  const v = _kenhSiRevenueOfMonthRaw_(monthKey);
  if (ck) _kpiCacheNumPut_(ck, v);
  return v;
}
function _kenhSiRevenueOfMonthRaw_(monthKey){
  const orders = _getOrdersForMonth(monthKey);
  return orders.reduce(function(a,o){ return a+o.total; }, 0);
}
// [VAT T11/2026] Doanh thu kênh sỉ CHƯA VAT = doanh thu có VAT ÷ 1,08
function _kenhSiRevenueNetOfMonth(monthKey){
  return _kpiNet_(_kenhSiRevenueOfMonth(monthKey));
}

function _calcKenhSiFund(monthKey){
  // [VAT T11/2026] từ T11/2026: Target = TB3T (chưa VAT) × hệ số; doanh thu tính % cũng chưa VAT. Tháng cũ: như cũ.
  const vat = _kpiVatOn_(monthKey);
  const valFn = vat ? _kenhSiRevenueNetOfMonth : _kenhSiRevenueOfMonth;
  const avgRes = _avgOverAvailableMonths(monthKey, valFn, 3);
  const heSo = _kpiHeSoThuong(monthKey);
  const target = avgRes.hasData ? Math.round(avgRes.avg*heSo) : null;
  const revThis = valFn(monthKey);
  const revThisGross = vat ? _kenhSiRevenueOfMonth(monthKey) : revThis;
  const isCurMonth = _isCurrentMonth(monthKey);
  const revProjected = isCurMonth ? _projectMetric(revThis, monthKey) : null;
  const revProjectedGross = isCurMonth ? (vat ? _projectMetric(revThisGross, monthKey) : revProjected) : null;
  const revForCalc = isCurMonth ? revProjected : revThis;

  let tyLeHoanThanh;
  if(target===null || target<=0){
    tyLeHoanThanh = revForCalc>0 ? 100 : 0;
  } else {
    tyLeHoanThanh = revForCalc/target*100;
  }

  let quyB2B=0, quyCSKH=0, tier='Dưới 80% - Không đạt (0đ)';
  if(tyLeHoanThanh>=100){ quyB2B=3000000; quyCSKH=5000000; tier='≥100%'; }
  else if(tyLeHoanThanh>=90){ quyB2B=2400000; quyCSKH=4000000; tier='90% - <100%'; }
  else if(tyLeHoanThanh>=80){ quyB2B=1500000; quyCSKH=2500000; tier='80% - <90%'; }

  return {
    // target / revThis / revProjected = số DÙNG ĐỂ TÍNH (từ T11/2026 là CHƯA VAT); *Gross = có VAT, *Net = chưa VAT
    target: target, revThis: revThis, revProjected: revProjected, isCurrentMonth: isCurMonth,
    tyLeHoanThanh: tyLeHoanThanh, tier: tier,
    quyB2B: quyB2B, quyCSKH: quyCSKH,
    hasEnoughData: avgRes.hasData, monthsUsed: avgRes.monthsUsed,
    vatApplied: vat, vatRate: KPI_VAT_RATE,
    revThisGross: revThisGross, revThisNet: vat ? revThis : null,
    revProjectedGross: revProjectedGross, revProjectedNet: vat ? revProjected : null
  };
}

function getKenhSiFundInfo(monthKey){
  try{ return _calcKenhSiFund(monthKey); } catch(e){ return {error:e.toString()}; }
}

// [MỚI #8] Chi tiết đơn hàng cấu thành doanh thu Kênh Sỉ của 1 tháng
function getKenhSiRevenueDetail(monthKey){
  try{
    const orders = _getOrdersForMonth(monthKey);
    return orders.map(function(o){
      return { date:o.date, name:o.name, phone:o.phone, emp:o.emp, total:o.total, totalNet:_kpiNet_(o.total), priceType:o.priceType };
    }).sort(function(a,b){ return new Date(b.date)-new Date(a.date); });
  } catch(e){ return []; }
}

function _calcNVB2BBonus(input, kenhSiPerPerson){
  const target=Number(input.targetNewCust)||0;
  const actualRaw=Number(input.actualNewCust)||0;
  const monthKey = input.monthKey;
  const isCurMonth = monthKey ? _isCurrentMonth(monthKey) : false;
  const actual = isCurMonth ? Math.round(_projectMetric(actualRaw, monthKey)) : actualRaw;
  let newCustBonus;
  if(target<=0){
    newCustBonus = actual>0 ? 900000 : 0;
  } else if(actual>=target){
    newCustBonus = 900000 + 200000*(actual-target);
  } else if(actual === target-1){
    newCustBonus = 450000; // thiếu đúng 1 khách so với mục tiêu → 50%
  } else {
    newCustBonus = 0;
  }

  const projectTarget=Number(input.projectTarget)||0;
  const projectCompleted=Number(input.projectCompleted)||0;
  const projectPct = projectTarget>0 ? (projectCompleted/projectTarget*100) : (projectCompleted>0?100:0);
  const projectBonus = projectPct>=100 ? 600000 : 0;

  const kenhSiBonus = Math.round(kenhSiPerPerson||0);

  const v1=Number(input.viol1)||0, v2=Number(input.viol2)||0, v3=Number(input.viol3)||0, v4=Number(input.viol4)||0;
  const totalViolations=v1+v2+v3+v4;
  const personalScore=Math.max(0, 100-10*totalViolations);
  const deduction=100000*totalViolations;

  const rawBonus=newCustBonus+projectBonus+kenhSiBonus;
  let finalB2BBonus, blocked=false;
  if(personalScore<60){ finalB2BBonus=0; blocked=true; }
  else { finalB2BBonus=Math.max(0, rawBonus-deduction); }

  const heSoNgayCong = _heSoNgayCong(input.gioCongThucTe);
  finalB2BBonus = Math.round(finalB2BBonus * heSoNgayCong);

  return { newCustBonus:newCustBonus,
    projectPct:projectPct, projectBonus:projectBonus,
    kenhSiBonus:kenhSiBonus,
    totalViolations:totalViolations, personalScore:personalScore,
    deduction:deduction, rawBonus:rawBonus, blocked:blocked, finalB2BBonus:finalB2BBonus,
    heSoNgayCong:heSoNgayCong,
    opsBonus:0, grandTotal:finalB2BBonus };
}

function _calcQLB2BBonus(input, kenhSiInfo){
  // Mục 1: tính theo TỔNG KH mới của các NVKD B2B. Target tổng = 2 KH × số NVKD B2B.
  //   đạt đủ target tổng -> 900.000đ | đạt từ 50% -> 450.000đ | dưới 50% -> 0đ
  const totalB2BCount = _getEmployeeListCached().filter(function(e){return e.role==='B2B';}).length || 1;
  const rawNewCust = Number(input.actualNewCust)||0;
  const isCurMonth = input.monthKey ? _isCurrentMonth(input.monthKey) : false;
  const qlActual = isCurMonth ? Math.round(_projectMetric(rawNewCust, input.monthKey)) : rawNewCust;
  const teamTarget = 2*totalB2BCount;
  let newCustBonus = 0;
  if(qlActual>=teamTarget) newCustBonus = 900000;
  else if(qlActual*2>=teamTarget) newCustBonus = 450000;

  const tyLe = kenhSiInfo ? kenhSiInfo.tyLeHoanThanh : 0;
  let serviceBonus=0, serviceTier='Dưới 80% - Không đạt (0đ)';
  if(tyLe>=100){ serviceBonus=1500000; serviceTier='≥100%'; }
  else if(tyLe>=90){ serviceBonus=1200000; serviceTier='90% - <100%'; }
  else if(tyLe>=80){ serviceBonus=750000; serviceTier='80% - <90%'; }

  const projectTarget=Number(input.projectTarget)||0;
  const projectCompleted=Number(input.projectCompleted)||0;
  const projectPct = projectTarget>0 ? (projectCompleted/projectTarget*100) : (projectCompleted>0?100:0);
  const projectBonus = projectPct>=100 ? 600000 : 0;

  // Mục 4: điểm hành vi TB của toàn bộ NVKD B2B + CSKH (tự động). >=80% -> 1.000.000đ; 50%-<80% -> 500.000đ; <50% -> 0đ
  const teamPct=Number(input.teamPct)||0;
  let opsBonus = 0;
  if(teamPct>=80) opsBonus = 1000000;
  else if(teamPct>=50) opsBonus = 500000;

  const v1=Number(input.viol1)||0, v2=Number(input.viol2)||0, v3=Number(input.viol3)||0, v4=Number(input.viol4)||0;
  const totalViolations=v1+v2+v3+v4;
  const personalScore=Math.max(0, 100-10*totalViolations);
  const deduction=100000*totalViolations;

  // Thưởng thực nhận = (Mục 1 + Mục 2 + Mục 3 + Mục 4 - Khấu trừ) × % Ngày công
  const rawBonus=newCustBonus+serviceBonus+projectBonus+opsBonus;
  let finalB2BBonus, blocked=false;
  if(personalScore<60){ finalB2BBonus=0; blocked=true; }
  else { finalB2BBonus=Math.max(0, rawBonus-deduction); }

  const heSoNgayCong = _heSoNgayCong(input.gioCongThucTe);
  finalB2BBonus = Math.round(finalB2BBonus * heSoNgayCong);

  return { newCustBonus:newCustBonus, teamTarget:teamTarget, teamActual:qlActual,
    servicePct:tyLe, serviceBonus:serviceBonus, serviceTier:serviceTier,
    projectPct:projectPct, projectBonus:projectBonus,
    totalViolations:totalViolations, personalScore:personalScore,
    deduction:deduction, rawBonus:rawBonus, blocked:blocked, finalB2BBonus:finalB2BBonus,
    heSoNgayCong:heSoNgayCong,
    opsBonus:opsBonus, teamPct:teamPct, grandTotal:finalB2BBonus };
}

function calcKPIBonusPreview(input){
  try{
    const monthKey = input.monthKey || _monthKey(new Date());
    const fund = _calcKenhSiFund(monthKey);
    if(input.isManager){
      const res = _calcQLB2BBonus(input, fund);
      res.kenhSiFund = fund;
      return res;
    } else {
      const nvCountB2B = _getEmployeeListCached().filter(function(e){return e.role==='B2B';}).length || 1;
      const kenhSiPerPerson = fund.quyB2B/nvCountB2B;
      const res = _calcNVB2BBonus(input, kenhSiPerPerson);
      res.kenhSiFund = fund;
      res.kenhSiPerPerson = kenhSiPerPerson;
      return res;
    }
  } catch(e){ return {error:e.toString()}; }
}

// [FIX #1] getKPIBonusList: QLB2B KH mới = SỐ NHÂN VIÊN B2B đã đạt đủ
// target cá nhân của họ trong tháng (không cộng dồn số khách thô).
// Mỗi nhân viên xử lý riêng trong try/catch - 1 người lỗi không làm
// sập cả danh sách.
function getKPIBonusList(monthKey, keepCache){
  try{
    if (!keepCache) { _KPI_ORDERS_CACHE = {}; _KPI_MT_TOTAL_CACHE = {}; _KPI_MT_NET_CACHE = {}; _KPI_EMP_LIST_CACHE = null; }

    const allEmp = _getEmployeeListCached();
    const empB2B = allEmp.filter(function(e){return e.role==='B2B';});
    const qlB2B = allEmp.filter(function(e){return e.role==='QLB2B';});
    const employees = empB2B.concat(qlB2B);
    if(!employees.length) return [];

    const sh=_kpiBonusSheetGioCong();
    const lastRow=sh.getLastRow();
    const savedRows = lastRow>1 ? sh.getRange(2,1,lastRow-1,KPI_COL_GIOCONG).getValues() : [];
    const savedMap={};
    savedRows.forEach(function(r){
      if(_kpiMonthKeyOf(r[0])===monthKey){
        savedMap[(r[1]||'').toString()]={
          targetNewCust:Number(r[3])||0, actualNewCust:Number(r[4])||0,
          projectTarget:Number(r[7])||0, projectCompleted:Number(r[8])||0,
          viol1:Number(r[9])||0, viol2:Number(r[10])||0, viol3:Number(r[11])||0, viol4:Number(r[12])||0,
          teamPct:Number(r[13])||0, note:r[14]||'',
          gioCongThucTe:Number(r[KPI_COL_GIOCONG-1])||208
        };
      }
    });

    const fund = _calcKenhSiFund(monthKey);
    const nvCountB2B = empB2B.length || 1;
    const kenhSiPerPersonB2B = fund.quyB2B/nvCountB2B;

    const mtGrowth = getMTGrowthForMonth(monthKey);
    const newCustMap=_newCustomersByEmployeeForMonth(monthKey);
    const isCurMonth = _isCurrentMonth(monthKey);

    let totalNewCustRaw = 0; // tổng KH mới (đủ điều kiện) của các NVKD B2B
    const teamPctAuto = _teamBehaviorAvgPct(monthKey); // điểm hành vi TB của B2B + CSKH (Mục 4 QLB2B)
    const empResults = [];
    empB2B.forEach(function(e){
      try{
        const saved=savedMap[e.displayName];
        const actualNewCustDefault = newCustMap[e.displayName]||0;
        const actualNewCustFinal = saved?saved.actualNewCust:actualNewCustDefault;
        const newCustProjected = isCurMonth ? _projectMetric(actualNewCustFinal, monthKey) : actualNewCustFinal;
        const targetNewCustThis = saved?saved.targetNewCust:2;
        totalNewCustRaw += actualNewCustFinal;

        const input = {
          monthKey: monthKey, isManager:false,
          targetNewCust: targetNewCustThis,
          actualNewCust: actualNewCustFinal,
          projectTarget: saved?saved.projectTarget:mtGrowth.mtTargetAvg3,
          projectCompleted: saved?saved.projectCompleted:(isCurMonth?mtGrowth.mtProjected:mtGrowth.mtThis),
          viol1:saved?saved.viol1:0, viol2:saved?saved.viol2:0, viol3:saved?saved.viol3:0, viol4:saved?saved.viol4:0,
          note: saved?saved.note:'',
          gioCongThucTe: saved?saved.gioCongThucTe:208
        };
        const calc=_calcNVB2BBonus(input, kenhSiPerPersonB2B);
        empResults.push({name:e.displayName, role:e.role, input:input, calc:calc, fund:fund, mtGrowth:mtGrowth,
          isCurrentMonth:isCurMonth, newCustActual:actualNewCustFinal, newCustProjected:newCustProjected, error:null});
      } catch(errEmp){
        empResults.push({name:e.displayName, role:e.role, error:errEmp.toString(),
          input:{targetNewCust:2,actualNewCust:0,projectTarget:0,projectCompleted:0,viol1:0,viol2:0,viol3:0,viol4:0,note:''},
          calc:{newCustBonus:0,projectPct:0,projectBonus:0,kenhSiBonus:0,totalViolations:0,personalScore:100,deduction:0,rawBonus:0,blocked:false,finalB2BBonus:0,opsBonus:0,grandTotal:0},
          fund:fund, mtGrowth:mtGrowth, isCurrentMonth:isCurMonth, newCustActual:0, newCustProjected:0});
      }
    });

    const results = empResults.slice();

    qlB2B.forEach(function(e){
      try{
        const saved=savedMap[e.displayName];
        const qlActualFinal = saved?saved.actualNewCust:totalNewCustRaw;
        const qlProjected = isCurMonth ? _projectMetric(qlActualFinal, monthKey) : qlActualFinal;
        const input = {
          monthKey: monthKey, isManager:true,
          targetNewCust: 2*(empB2B.length||1),
          actualNewCust: qlActualFinal,
          projectTarget: saved?saved.projectTarget:mtGrowth.mtTargetAvg3,
          projectCompleted: saved?saved.projectCompleted:(isCurMonth?mtGrowth.mtProjected:mtGrowth.mtThis),
          viol1:saved?saved.viol1:0, viol2:saved?saved.viol2:0, viol3:saved?saved.viol3:0, viol4:saved?saved.viol4:0,
          teamPct: teamPctAuto,
          note: saved?saved.note:'',
          gioCongThucTe: saved?saved.gioCongThucTe:208
        };
        const calc=_calcQLB2BBonus(input, fund);
        results.push({name:e.displayName, role:e.role, input:input, calc:calc, fund:fund, mtGrowth:mtGrowth,
          isCurrentMonth:isCurMonth, newCustActual:qlActualFinal, newCustProjected:qlProjected,
          totalNewCustB2B:totalNewCustRaw, totalB2BCount:empB2B.length, error:null});
      } catch(errEmp){
        results.push({name:e.displayName, role:e.role, error:errEmp.toString(),
          input:{targetNewCust:2,actualNewCust:0,projectTarget:0,projectCompleted:0,viol1:0,viol2:0,viol3:0,viol4:0,teamPct:0,note:''},
          calc:{newCustBonus:0,servicePct:0,serviceBonus:0,serviceTier:'',projectPct:0,projectBonus:0,totalViolations:0,personalScore:100,deduction:0,rawBonus:0,blocked:false,finalB2BBonus:0,opsBonus:0,teamPct:0,grandTotal:0},
          fund:fund, mtGrowth:mtGrowth, isCurrentMonth:isCurMonth, newCustActual:0, newCustProjected:0,
          totalNewCustB2B:totalNewCustRaw, totalB2BCount:empB2B.length});
      }
    });

    return results.sort(function(a,b){ return (b.calc.grandTotal||0)-(a.calc.grandTotal||0); });
  } catch(e){ return { error: e.toString() }; }
}

function getTargetDetailBreakdown(monthKey){
  try{
    const cfg = (typeof getThangCaoDiem === 'function') ? getThangCaoDiem(monthKey) : null;
    const heSoThuong = cfg ? (1+cfg.pctTang/100) : 1.10;
    const heSoMT = cfg ? (1+cfg.pctTang/100) : 1.20;

    // [VAT T11/2026] từ T11/2026: giá trị 3 tháng dùng để tính Target là CHƯA VAT (cột "value"); cột "valueGross" là có VAT
    const vat = _kpiVatOn_(monthKey);
    const kenhSiDetail = _avgOverAvailableMonthsDetail(monthKey, vat ? _kenhSiRevenueNetOfMonth : _kenhSiRevenueOfMonth, 3);
    const mtDetail = _avgOverAvailableMonthsDetail(monthKey, vat ? _mtNetRevenueOfMonth : _mtTotalRevenueOfMonth, 3);

    function withTarget(detail, heSo, unitSuffix, grossFn){
      const target = detail.hasData ? Math.round(detail.avg*heSo) : null;
      const months = detail.months.map(function(m){
        const p=m.monthKey.split('-');
        return { monthKey:m.monthKey, label:'Th.'+p[1]+'/'+p[0], value:m.value,
                 valueGross: vat ? grossFn(m.monthKey) : m.value, valueNet: vat ? m.value : null };
      });
      let avgGross = null;
      if (vat) {
        const gv = months.map(function(x){ return x.valueGross; }).filter(function(v){ return v>0; });
        avgGross = gv.length ? gv.reduce(function(a,b){return a+b;},0)/gv.length : null;
      }
      return {
        months: months,
        avg: detail.avg, avgGross: avgGross, monthsUsed: detail.monthsUsed, hasData: detail.hasData,
        heSo: heSo, heSoLabel: (heSo*100).toFixed(0)+'%',
        isSpecialMonth: !!cfg, pctTang: cfg?cfg.pctTang:0, lyDo: cfg?cfg.lyDo:'',
        target: target, unitSuffix: unitSuffix||'',
        vatApplied: vat, vatRate: KPI_VAT_RATE
      };
    }

    return {
      monthKey: monthKey, vatApplied: vat, vatRate: KPI_VAT_RATE,
      kenhSi: withTarget(kenhSiDetail, heSoThuong, 'đ', _kenhSiRevenueOfMonth),
      mtGrowth: withTarget(mtDetail, heSoMT, 'đ', _mtTotalRevenueOfMonth)
    };
  } catch(e){ return { error: e.toString() }; }
}

function repairKPIThuongMonthColumn(){
  const sh = _kpiBonusSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log('Sheet KPIThuong trống.'); return; }
  const data = sh.getRange(2, 1, lastRow - 1, 16).getValues();
  const map = {};
  data.forEach(function(r, i){
    const mk = _kpiMonthKeyOf(r[0]);
    const emp = (r[1]||'').toString();
    const key = mk + '|' + emp;
    const updatedAt = r[15] instanceof Date ? r[15].getTime() : 0;
    if (!map[key] || updatedAt >= map[key].updatedAt) {
      map[key] = { updatedAt: updatedAt, row: r.slice() };
      map[key].row[0] = mk;
    }
  });
  const keptRows = Object.keys(map).map(function(k){ return map[k].row; });
  sh.getRange(2, 1, lastRow - 1, 16).clearContent();
  if (keptRows.length) {
    sh.getRange(2, 1, keptRows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, keptRows.length, 16).setValues(keptRows);
  }
  Logger.log('Đã dọn KPIThuong: giữ lại ' + keptRows.length + ' dòng.');
}
// ═══ TĂNG TỐC: cache + gom 1 lượt gọi ═══
var _KPI_VER_MEM = null;
function _kpiVer_(){
  if (_KPI_VER_MEM) return _KPI_VER_MEM;
  try { const c = CacheService.getScriptCache(); let v = c.get('KPI_VER'); if (!v) { v = String(Date.now()); c.put('KPI_VER', v, 21600); } _KPI_VER_MEM = v; return v; }
  catch (e) { return '0'; }
}
// Gọi hàm này sau khi duyệt/sửa/xóa đơn nếu muốn số KPI cập nhật ngay (không chờ 5 phút)
function _kpiBust_(){ try { CacheService.getScriptCache().put('KPI_VER', String(Date.now()), 21600); } catch (e) {} _KPI_VER_MEM = null; }
function _kpiPastKey_(prefix, monthKey){ return (monthKey < _monthKey(new Date())) ? ('KPI_' + prefix + '_' + _kpiVer_() + '_' + monthKey) : ''; }
function _kpiCacheNumGet_(k){ try { const v = CacheService.getScriptCache().get(k); return v === null ? null : Number(v); } catch (e) { return null; } }
function _kpiCacheNumPut_(k, n){ try { CacheService.getScriptCache().put(k, String(n), 300); } catch (e) {} }
function _kpiBigPut_(key, obj, ttl){
  try {
    const s = JSON.stringify(obj), n = Math.ceil(s.length / 90000), m = {};
    m[key + '_n'] = String(n);
    for (let i = 0; i < n; i++) m[key + '_' + i] = s.substr(i * 90000, 90000);
    CacheService.getScriptCache().putAll(m, ttl);
  } catch (e) {}
}
function _kpiBigGet_(key){
  try {
    const c = CacheService.getScriptCache(), n = Number(c.get(key + '_n'));
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const m = c.getAll(keys); let s = '';
    for (let i = 0; i < n; i++) { const p = m[key + '_' + i]; if (p == null) return null; s += p; }
    return JSON.parse(s);
  } catch (e) { return null; }
}
var _KPI_FO_FRESH = false;
function _kpiFirstOrderMapCached_(force){
  const key = 'KPI_FO_' + _kpiVer_();
  if (!force) {
    const c = _kpiBigGet_(key);
    if (c) { const m = {}; Object.keys(c).forEach(function(k){ const v = c[k]; m[k] = (v && typeof v === 'object' && v.t !== undefined) ? new Date(v.t) : v; }); return m; }
  }
  const src = _loadCustomerFirstOrderMap();
  _KPI_FO_FRESH = true;
  const out = {};
  Object.keys(src).forEach(function(k){ const v = src[k]; out[k] = (v instanceof Date) ? { t: v.getTime() } : v; });
  _kpiBigPut_(key, out, 300);
  return src;
}
// 1 lượt gọi duy nhất cho cả trang Thưởng KPI: 4 phần dùng chung bộ nhớ nên không đọc lại sheet
function getKPIBootstrap(monthKey){
  try {
    _KPI_ORDERS_CACHE = {}; _KPI_MT_TOTAL_CACHE = {}; _KPI_MT_NET_CACHE = {}; _KPI_EMP_LIST_CACHE = null;
    const out = { monthKey: monthKey };
    out.list = getKPIBonusList(monthKey, true);
    try { out.cskh = getKPIBonusListCSKH(monthKey); } catch (e1) { out.cskh = { error: e1.toString() }; }
    try { out.target = getTargetSummary(monthKey); } catch (e2) { out.target = { error: e2.toString() }; }
    out.mt = getMTGrowthForMonth(monthKey);
    return out;
  } catch (e) { return { error: e.toString() }; }
}