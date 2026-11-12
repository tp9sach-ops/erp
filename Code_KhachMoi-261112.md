/** ============================================================
 *  9 SẠCH - KHÁCH MỚI CHỜ DUYỆT  (FILE MỚI)
 *  Luồng: khách chưa có code -> hệ thống tự cấp mã -> popup xác nhận -> lưu vào sheet
 *  "KhachMoiChoDuyet" (chưa vào DataKhachSi) -> (KHÔNG gửi Zalo khi tạo) ->
 *  Admin/QLB2B bấm Duyệt -> khách chính thức vào DataKhachSi (giữ nguyên mã đã cấp).
 *
 *  !!! TRƯỚC KHI DÁN: dùng Ctrl+F tìm "createPendingCustomerInternal_" trong toàn project.
 *      Nếu đã có hàm này ở file khác thì KHÔNG dán file này (trùng tên sẽ chạy sai).
 *  Dùng lại: _sheet, _dataKhachSiSheet, DKS_COL, DKS_TOTAL_COLS, _generateCode_, _readAllCodes_,
 *            _normCode_, _stripDiacriticsServer (Code.gs); guiZaloTheoVaiTro_, guiZaloTheoTenNV_ (Notification.gs).
 * ============================================================ */

const PEND_SHEET = 'KhachMoiChoDuyet';
const PEND_HEADERS = ['Mã KH', 'Tên', 'SĐT', 'Địa chỉ', 'Tỉnh/Thành', 'Nhóm', 'Xưởng', 'Mã Tỉnh', 'MST', 'Công ty',
  'Tên Xuất HĐ', 'CCCD', 'Loại giá', 'NV tạo', 'Ngày tạo', 'Trạng thái', 'Người duyệt', 'Vai trò duyệt', 'Ngày duyệt', 'Lý do'];
const PC = { CODE:1, NAME:2, PHONE:3, ADDR:4, CITY:5, GROUP:6, FACTORY:7, PROV:8, TAX:9, COMPANY:10, INVNAME:11, CCCD:12,
  PRICE:13, EMP:14, CREATED:15, STATUS:16, BY:17, BYROLE:18, AT:19, REASON:20 };
const PEND_APPROVER_ROLES = ['Admin', 'QLB2B'];

function _pendSheet_() {
  const sh = _sheet(PEND_SHEET, PEND_HEADERS);
  sh.getRange(2, PC.PHONE, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  sh.getRange(2, PC.CODE, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  sh.getRange(2, PC.TAX, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  sh.getRange(2, PC.CCCD, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  return sh;
}
function _pendRows_() {
  const sh = _pendSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, PEND_HEADERS.length).getValues().map(function (r, i) { return { row: i + 2, v: r }; });
}
function _pendFmt_(d) {
  if (!d) return '';
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'HH:mm dd/MM/yyyy'); } catch (e) { return String(d); }
}
function _pendObj_(x) {
  const v = x.v;
  return {
    sheetRow: x.row, code: _normCode_(v[PC.CODE - 1]), name: String(v[PC.NAME - 1] || ''), phone: String(v[PC.PHONE - 1] || ''),
    address: String(v[PC.ADDR - 1] || ''), city: String(v[PC.CITY - 1] || ''), group: String(v[PC.GROUP - 1] || ''),
    factory: String(v[PC.FACTORY - 1] || ''), prov: String(v[PC.PROV - 1] || ''), taxCode: String(v[PC.TAX - 1] || ''),
    company: String(v[PC.COMPANY - 1] || ''), invoiceName: String(v[PC.INVNAME - 1] || ''), cccd: String(v[PC.CCCD - 1] || ''),
    customerType: String(v[PC.PRICE - 1] || ''), emp: String(v[PC.EMP - 1] || ''), createdAt: _pendFmt_(v[PC.CREATED - 1]),
    status: String(v[PC.STATUS - 1] || ''), approvedBy: String(v[PC.BY - 1] || ''), approvedRole: String(v[PC.BYROLE - 1] || ''),
    approvedAt: _pendFmt_(v[PC.AT - 1]), reason: String(v[PC.REASON - 1] || '')
  };
}

// Mã của MỌI khách trong sheet chờ duyệt (kể cả bị từ chối) -> không cấp trùng
function _pendCodes_() { return _pendRows_().map(function (x) { return x.v[PC.CODE - 1]; }); }

function _pendFindByCode_(code) {
  const c = _normCode_(code);
  if (!c) return null;
  const f = _pendRows_().filter(function (x) { return _normCode_(x.v[PC.CODE - 1]) === c && x.v[PC.STATUS - 1] === 'Chờ duyệt'; })[0];
  return f ? _pendObj_(f) : null;
}

// Gợi ý khách chờ duyệt trong ô tìm kiếm khi lên đơn (đánh dấu pending)
function _pendSearch_(query) {
  const q = _stripDiacriticsServer(query || '').trim();
  if (!q) return [];
  const toks = q.split(/\s+/).filter(Boolean);
  const out = [];
  _pendRows_().forEach(function (x) {
    if (x.v[PC.STATUS - 1] !== 'Chờ duyệt') return;
    const o = _pendObj_(x);
    const hay = _stripDiacriticsServer([o.name, o.code, o.phone, o.address, o.company, o.taxCode].join(' '));
    if (toks.every(function (t) { return hay.indexOf(t) !== -1; })) {
      o.pending = true;
      out.push(o);
    }
  });
  return out.slice(0, 5);
}

// Gọi từ submitOrder (đang giữ lock). Trả {error} | {pending:true, code} | {pending:false, code}
function createPendingCustomerInternal_(data) {
  const phone = (data.phone || '').toString().replace(/[^0-9]/g, '');
  if (!data.name || phone.length < 9) return { error: 'Khách mới cần có Tên và SĐT (tối thiểu 9 số).' };

  // 1) Đã là khách chính thức có mã (trùng SĐT) -> dùng mã sẵn có
  const dks = _dataKhachSiSheet();
  if (dks.getLastRow() > 1) {
    const rows = dks.getRange(2, 1, dks.getLastRow() - 1, DKS_TOTAL_COLS).getValues();
    for (let i = 0; i < rows.length; i++) {
      const p = (rows[i][DKS_COL.PHONE - 1] || '').toString().replace(/[^0-9]/g, '');
      const c = _normCode_(rows[i][DKS_COL.CODE - 1]);
      if (p === phone && c) return { pending: false, code: c };
    }
  }
  // 2) Đã có hồ sơ chờ duyệt cùng SĐT -> dùng lại mã đó (không tạo trùng)
  const ex = _pendRows_().filter(function (x) {
    return String(x.v[PC.PHONE - 1] || '').replace(/[^0-9]/g, '') === phone && x.v[PC.STATUS - 1] === 'Chờ duyệt';
  })[0];
  if (ex) return { pending: true, code: _normCode_(ex.v[PC.CODE - 1]) };

  // 3) Cấp mã mới theo quy định + lưu chờ duyệt
  const gen = _generateCode_(_readAllCodes_(dks), data.codeGroup, data.codeFactory, data.city || '', data.address || '');
  if (!gen.code) return { error: gen.error || 'Không tạo được mã khách mới.' };
  const sh = _pendSheet_();
  const arr = new Array(PEND_HEADERS.length).fill('');
  arr[PC.CODE - 1] = gen.code; arr[PC.NAME - 1] = data.name; arr[PC.PHONE - 1] = phone; arr[PC.ADDR - 1] = data.address || '';
  arr[PC.CITY - 1] = data.city || ''; arr[PC.GROUP - 1] = gen.group; arr[PC.FACTORY - 1] = gen.factory; arr[PC.PROV - 1] = gen.prov;
  arr[PC.TAX - 1] = data.taxCode || ''; arr[PC.COMPANY - 1] = data.company || ''; arr[PC.INVNAME - 1] = data.invoiceName || '';
  arr[PC.CCCD - 1] = data.cccd || ''; arr[PC.PRICE - 1] = data.priceType || ''; arr[PC.EMP - 1] = data.employeeName || '';
  arr[PC.CREATED - 1] = new Date(); arr[PC.STATUS - 1] = 'Chờ duyệt';
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, arr.length).setValues([arr]);

  // [ĐÃ BỎ] Không gửi Zalo khi tạo khách mới - Admin/QLB2B tự vào Data Khách Sỉ > Khách mới chờ duyệt để duyệt.
  return { pending: true, code: gen.code };
}

function _pendRoleOf_(username) {
  const sh = _sheet('NhanVien');
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  const u = String(username || '').trim();
  const f = rows.filter(function (r) { return String(r[1]).trim() === u; })[0];
  return f ? { name: String(f[0]).trim(), role: String(f[3] || 'NhanVien').trim() } : null;
}

// [Client gọi] Danh sách khách mới. Admin/QLB2B thấy hết; NV chỉ thấy khách do mình tạo.
function getPendingCustomers(username) {
  try {
    const me = _pendRoleOf_(username);
    if (!me) return { error: 'Không xác định được tài khoản.' };
    const can = PEND_APPROVER_ROLES.indexOf(me.role) !== -1;
    let list = _pendRows_().map(_pendObj_);
    if (!can) list = list.filter(function (x) { return x.emp === me.name; });
    list.sort(function (a, b) { return (a.status === 'Chờ duyệt' ? 0 : 1) - (b.status === 'Chờ duyệt' ? 0 : 1) || b.sheetRow - a.sheetRow; });
    return { list: list, canApprove: can };
  } catch (e) { return { error: e.toString() }; }
}

// [Client gọi] Duyệt -> chính thức vào DataKhachSi (giữ đúng mã đã cấp)
function approvePendingCustomer(username, sheetRow) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    const me = _pendRoleOf_(username);
    if (!me || PEND_APPROVER_ROLES.indexOf(me.role) === -1) return { success: false, error: 'Chỉ Admin / QLB2B được duyệt.' };
    const sh = _pendSheet_();
    const row = parseInt(sheetRow, 10);
    if (!row || row < 2 || row > sh.getLastRow()) return { success: false, error: 'Dòng không hợp lệ.' };
    const v = sh.getRange(row, 1, 1, PEND_HEADERS.length).getValues()[0];
    if (v[PC.STATUS - 1] !== 'Chờ duyệt') return { success: false, error: 'Khách này đã được xử lý (' + v[PC.STATUS - 1] + ').' };
    const code = _normCode_(v[PC.CODE - 1]);
    if (_readAllCodesOfficial_().indexOf(code) !== -1) return { success: false, error: 'Mã ' + code + ' đã tồn tại trong Data Khách Sỉ.' };

    const dks = _dataKhachSiSheet();
    const phone = String(v[PC.PHONE - 1] || '').replace(/[^0-9]/g, '');
    const today = new Date();
    let target = -1;
    if (dks.getLastRow() > 1) {                       // trùng SĐT với khách cũ chưa có mã -> gắn mã vào dòng đó
      const rows = dks.getRange(2, 1, dks.getLastRow() - 1, DKS_TOTAL_COLS).getValues();
      for (let i = 0; i < rows.length; i++) {
        const p = String(rows[i][DKS_COL.PHONE - 1] || '').replace(/[^0-9]/g, '');
        if (p === phone && !_normCode_(rows[i][DKS_COL.CODE - 1])) { target = i + 2; break; }
      }
    }
    const arr = (target > 0) ? dks.getRange(target, 1, 1, DKS_TOTAL_COLS).getValues()[0] : new Array(DKS_TOTAL_COLS).fill('');
    arr[DKS_COL.NAME - 1] = v[PC.NAME - 1]; arr[DKS_COL.PHONE - 1] = phone; arr[DKS_COL.ADDRESS - 1] = v[PC.ADDR - 1];
    if (!arr[DKS_COL.FIRST_ORDER - 1]) arr[DKS_COL.FIRST_ORDER - 1] = v[PC.CREATED - 1] || today;
    arr[DKS_COL.LAST_ORDER - 1] = today;
    if (v[PC.PRICE - 1]) arr[DKS_COL.LOAI_KHACH_SI - 1] = v[PC.PRICE - 1];
    if (v[PC.TAX - 1]) arr[DKS_COL.TAX - 1] = v[PC.TAX - 1];
    if (v[PC.COMPANY - 1]) arr[DKS_COL.COMPANY - 1] = v[PC.COMPANY - 1];
    if (v[PC.INVNAME - 1]) arr[DKS_COL.INVOICE_NAME - 1] = v[PC.INVNAME - 1];
    if (v[PC.CCCD - 1]) arr[DKS_COL.CCCD - 1] = v[PC.CCCD - 1];
    arr[DKS_COL.CODE - 1] = code; arr[DKS_COL.FACTORY - 1] = v[PC.FACTORY - 1];
    arr[DKS_COL.PROV - 1] = v[PC.PROV - 1]; arr[DKS_COL.GROUP - 1] = v[PC.GROUP - 1];
    const r1 = target > 0 ? target : dks.getLastRow() + 1;
    [DKS_COL.PHONE, DKS_COL.TAX, DKS_COL.CCCD, DKS_COL.CODE].forEach(function (c) { dks.getRange(r1, c).setNumberFormat('@'); });
    dks.getRange(r1, 1, 1, DKS_TOTAL_COLS).setValues([arr]);

    sh.getRange(row, PC.STATUS, 1, 4).setValues([['Đã duyệt', me.name, me.role, new Date()]]);
    // [ĐÃ BỎ] Không gửi Zalo khi duyệt khách mới.
    return { success: true, code: code };
  } catch (e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}

function _readAllCodesOfficial_() {
  const sh = _dataKhachSiSheet();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, DKS_COL.CODE, sh.getLastRow() - 1, 1).getValues().map(function (r) { return _normCode_(r[0]); });
}

// [Client gọi] Từ chối
function rejectPendingCustomer(username, sheetRow, reason) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, error: 'Hệ thống đang bận, thử lại sau.' }; }
  try {
    const me = _pendRoleOf_(username);
    if (!me || PEND_APPROVER_ROLES.indexOf(me.role) === -1) return { success: false, error: 'Chỉ Admin / QLB2B được từ chối.' };
    const sh = _pendSheet_();
    const row = parseInt(sheetRow, 10);
    if (!row || row < 2 || row > sh.getLastRow()) return { success: false, error: 'Dòng không hợp lệ.' };
    if (sh.getRange(row, PC.STATUS).getValue() !== 'Chờ duyệt') return { success: false, error: 'Khách này đã được xử lý.' };
    sh.getRange(row, PC.STATUS, 1, 4).setValues([['Từ chối', me.name, me.role, new Date()]]);
    sh.getRange(row, PC.REASON).setValue(reason || '');
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
  finally { lock.releaseLock(); }
}