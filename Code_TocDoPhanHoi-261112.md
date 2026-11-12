/** ============================================================
 *  9 SẠCH - TỐC ĐỘ PHẢN HỒI CSKH (file thống kê nhân viên Fanpage/Zalo)
 *  ================================================================
 *  CÁCH DÙNG: Apps Script Editor -> "+" cạnh Files -> Script -> đặt tên
 *  "Code_TocDoPhanHoi" -> dán TOÀN BỘ nội dung này. Dùng lại _sheet(),
 *  _clean(), _monthKey() có sẵn trong Code.gs.
 *
 *  Nguồn: file Excel "user_statistics_multi_pages_ALL_*.xlsx" (sheet
 *  "All By Staff"). Client đọc file, lấy dòng "Tất cả trang" của từng
 *  nhân viên (thời gian phản hồi TB toàn bộ các trang) rồi gửi lên đây.
 *  Mỗi (Tháng + Nhân viên) chỉ giữ 1 dòng - upload lại sẽ GHI ĐÈ.
 * ============================================================ */
const RS_HEADERS = ['Batch ID','Tháng','Tên Ảo','Tên Thật','Thời Gian Phản Hồi TB (giây)','Thời Gian (chữ)','SĐT Mang Về','Số Đơn Chốt','Tỷ Lệ Phản Hồi (%)','Từ Ngày','Đến Ngày','Cập Nhật Lúc','Nhân Viên Upload'];
function _rsSheet(){
  const sh = _sheet('TocDoPhanHoi', RS_HEADERS);
  try { sh.getRange(2, 2, Math.max(sh.getLastRow()-1,1), 1).setNumberFormat('@'); } catch(e){}
  return sh;
}
const RS_LOG_HEADERS = ['Batch ID','Thời Gian Upload','Tên File','Tháng','Số Nhân Viên','Nhân Viên Upload'];
function _rsLogSheet(){ return _sheet('TocDoPhanHoi_UploadLog', RS_LOG_HEADERS); }

// rowsArr: [{virtualName, realName, avgSeconds, avgText, sdt, orders, replyRate}]
// meta: {monthKey, fromDate, toDate, fileName, employeeName}
function importResponseSpeedBatch(rowsArr, meta){
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success:false, error:'Hệ thống đang bận, thử lại sau.' }; }
  try{
    const monthKey = (meta && meta.monthKey || '').toString().trim();
    if(!/^\d{4}-\d{2}$/.test(monthKey)) return { success:false, error:'Không xác định được tháng của file.' };
    const sh = _rsSheet();
    const batchId = 'RS' + new Date().getTime();
    const lastRow = sh.getLastRow();
    const existing = lastRow>1 ? sh.getRange(2,1,lastRow-1,RS_HEADERS.length).getValues() : [];
    const idx = {};
    existing.forEach(function(r,i){ idx[(r[1]||'').toString()+'|'+(r[3]||'').toString().trim()] = i+2; });
    let added=0, updated=0; const newRows=[];
    (rowsArr||[]).forEach(function(r){
      const real = (r.realName||'').toString().trim();
      if(!real) return;
      const row = [batchId, monthKey, r.virtualName||'', real, Number(r.avgSeconds)||0, r.avgText||'',
        Number(r.sdt)||0, Number(r.orders)||0, Number(r.replyRate)||0,
        meta.fromDate?new Date(meta.fromDate):'', meta.toDate?new Date(meta.toDate):'', new Date(), meta.employeeName||''];
      const key = monthKey+'|'+real;
      if(idx[key]){ sh.getRange(idx[key],1,1,row.length).setValues([row]); updated++; }
      else { newRows.push(row); added++; }
    });
    if(newRows.length){
      const start = sh.getLastRow()+1;
      sh.getRange(start,2,newRows.length,1).setNumberFormat('@');
      sh.getRange(start,1,newRows.length,RS_HEADERS.length).setValues(newRows);
    }
    _rsLogSheet().appendRow([batchId, new Date(), meta.fileName||'', monthKey, added+updated, meta.employeeName||'']);
    return { success:true, added:added, updated:updated, batchId:batchId, monthKey:monthKey };
  } catch(e){ return { success:false, error:e.toString() }; }
  finally { lock.releaseLock(); }
}

function getResponseSpeedUploadLog(){
  try{
    const sh = _rsLogSheet(); const lastRow = sh.getLastRow();
    if(lastRow<2) return [];
    return sh.getRange(2,1,lastRow-1,6).getValues().map(function(r){
      return { batchId:(r[0]||'').toString(), time:_clean(r[1]), fileName:r[2]||'', monthKey:(r[3]||'').toString(), count:Number(r[4])||0, emp:r[5]||'' };
    }).sort(function(a,b){ return new Date(b.time)-new Date(a.time); });
  } catch(e){ return []; }
}

// Xóa 1 lượt upload: chỉ xóa các dòng mà "Batch ID" còn đúng bằng batchId này
function deleteResponseSpeedBatch(batchId){
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success:false, error:'Hệ thống đang bận, thử lại sau.' }; }
  try{
    const sh = _rsSheet(); const lastRow = sh.getLastRow();
    if(lastRow>1){
      const data = sh.getRange(2,1,lastRow-1,1).getValues();
      const del=[]; data.forEach(function(r,i){ if((r[0]||'').toString()===batchId) del.push(i+2); });
      del.sort(function(a,b){return b-a;}).forEach(function(n){ sh.deleteRow(n); });
    }
    const lg = _rsLogSheet(); const ll = lg.getLastRow();
    if(ll>1){
      const d = lg.getRange(2,1,ll-1,1).getValues();
      for(let i=0;i<d.length;i++){ if((d[i][0]||'').toString()===batchId){ lg.deleteRow(i+2); break; } }
    }
    return { success:true };
  } catch(e){ return { success:false, error:e.toString() }; }
  finally { lock.releaseLock(); }
}

function _rsFmtText(sec){
  sec = Math.round(Number(sec)||0);
  return Math.floor(sec/60)+'m '+(sec%60)+'s';
}
// { tênThật: {avgSeconds, avgText, sdt, orders} } của 1 tháng
function getResponseSpeedMap(monthKey){
  const map = {};
  try{
    const sh = _rsSheet(); const lastRow = sh.getLastRow();
    if(lastRow<2) return map;
    sh.getRange(2,1,lastRow-1,RS_HEADERS.length).getValues().forEach(function(r){
      if((r[1]||'').toString()!==monthKey) return;
      const name=(r[3]||'').toString().trim(); if(!name) return;
      map[name] = { avgSeconds:Number(r[4])||0, avgText:(r[5]||'').toString()||_rsFmtText(r[4]), sdt:Number(r[6])||0, orders:Number(r[7])||0 };
    });
  } catch(e){}
  return map;
}
function getResponseSpeedForEmployee(monthKey, empName){
  return getResponseSpeedMap(monthKey)[(empName||'').toString().trim()] || null;
}