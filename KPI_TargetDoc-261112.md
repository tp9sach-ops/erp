/** ============================================================
 *  9 SẠCH - VĂN BẢN THÔNG BÁO TARGET KPI  (FILE MỚI - thêm vào project Apps Script)
 *  Tiêu đề có "TOÀN KÊNH B2B" để xác định văn bản CHỈ áp dụng cho kênh B2B.
 *  Nội dung cập nhật theo chính sách mới nhất của Quản lý B2B (QLB2B - quỹ tối đa 4.000.000đ: Mục 1 gồm 1a+1b, Mục 3 = 500.000đ, Mục 4 = 4 mốc 500/350/250/0) và Trưởng phòng Kinh doanh (TPKD).
 *  Số liệu lấy từ đúng các hàm tính đang dùng: getTargetDetailBreakdown, _calcKenhSiFund,
 *  getMTGrowthForMonth, getKPIBonusList -> không tính lại công thức ở nơi khác.
 *
 *  !!! BẮT BUỘC: XÓA hàm exportTargetAnnouncementDoc CŨ (ở file Apps Script cũ), vì file này
 *      định nghĩa lại hàm cùng tên (trùng tên sẽ chạy sai).
 *  Cần có sẵn: _appendKPIDocHeader + _kpiFinalizeDoc_ (file KPI_DocStyle.gs BẢN MỚI - có tham số opts).
 * ============================================================ */

function _tgtExportDocxB64_(doc) {
  if (typeof _exportDocAsBase64Docx === 'function') return _exportDocAsBase64Docx(doc);
  doc.saveAndClose();
  const url = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + doc.getId() + '&exportFormat=docx';
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Không xuất được Word (HTTP ' + resp.getResponseCode() + ').');
  const b64 = Utilities.base64Encode(resp.getBlob().getBytes());
  try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e) {}
  return b64;
}

function exportTargetAnnouncementDoc(monthKey, useActualOnly) {
  _KPI_FORCE_ACTUAL_ONLY = !!useActualOnly;
  let doc = null;
  try {
    if (!/^\d{4}-\d{2}$/.test(monthKey || '')) return { success: false, error: 'Tháng không hợp lệ.' };
    const parts = monthKey.split('-'), yyyy = parts[0], mm = parts[1];
    const now = new Date();
    const ngayStr = 'ngày ' + now.getDate() + ' tháng ' + (now.getMonth() + 1) + ' năm ' + now.getFullYear();
    const money = function (v) { return Math.round(v || 0).toLocaleString('vi-VN') + 'đ'; };
    const BLACK = '#000000';

    // ── số liệu (đúng các hàm đang tính KPI, không tính lại công thức) ──
    const det = getTargetDetailBreakdown(monthKey);
    if (det && det.error) return { success: false, error: det.error };
    const fund = _calcKenhSiFund(monthKey);
    const mt = getMTGrowthForMonth(monthKey);
    const list = getKPIBonusList(monthKey);
    if (list && list.error) return { success: false, error: list.error };
    const nvs = (list || []).filter(function (e) { return e.role === 'B2B' && !e.error; });
    const nB2B = nvs.length || 1;
    const special = det.kenhSi && det.kenhSi.isSpecialMonth;
    // [VAT T11/2026] từ T11/2026: Target và doanh thu tính % đều CHƯA VAT
    const vatOn = _kpiVatOn_(monthKey);
    const vatTag = vatOn ? ' (chưa VAT)' : '';
    const ksT = det.kenhSi && det.kenhSi.hasData ? (money(det.kenhSi.target) + vatTag) : 'Chưa đủ dữ liệu';
    const mtT = det.mtGrowth && det.mtGrowth.hasData ? (money(det.mtGrowth.target) + vatTag) : 'Chưa đủ dữ liệu';
    const nOk = nvs.filter(function (e) {
      const tg = Number(e.input.targetNewCust) || 2;
      const ac = Math.floor((e.isCurrentMonth && !useActualOnly) ? (e.newCustProjected || 0) : (e.newCustActual || 0));
      return ac >= tg;
    }).length;
    const m2aNow = nvs.length ? Math.round(500000 * nOk / nvs.length / 1000) * 1000 : 0;

    doc = DocumentApp.create('TEMP_TARGET_B2B_' + monthKey + '_' + new Date().getTime());
    const body = doc.getBody();
    // A4 + lề 55pt (~1,94cm) như văn bản cũ -> bảng không tràn lề
    body.setPageWidth(595.3).setPageHeight(841.9);
    body.setMarginTop(45).setMarginBottom(45).setMarginLeft(55).setMarginRight(55);
    const avail = 595.3 - 110;

    _appendKPIDocHeader(body, 'TARGET-B2B-' + mm + yyyy, ngayStr);

    // Tiêu đề 2 dòng, chữ lớn (20pt do _kpiFinalizeDoc_ chuẩn hóa)
    body.appendParagraph('THÔNG BÁO TARGET & MỨC THƯỞNG KPI').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('THÁNG ' + mm + '/' + yyyy + ' - TOÀN KÊNH B2B').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const pS = body.appendParagraph('Phạm vi áp dụng: CHỈ KÊNH B2B (NVKD B2B, CSKH kênh sỉ, Quản lý B2B, Trưởng phòng Kinh doanh phần B2B). Không áp dụng cho kênh B2C.');
    pS.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pS.editAsText().setBold(true).setFontSize(10).setForegroundColor(BLACK);
    body.appendParagraph(useActualOnly
      ? '📌 Số liệu trong văn bản này tính theo LŨY KẾ THỰC TẾ đến thời điểm xuất (không có phần dự kiến).'
      : '📌 Số liệu thực hiện của tháng hiện tại có bao gồm phần DỰ KIẾN hoàn thành đến cuối tháng.')
      .editAsText().setFontSize(9.5).setForegroundColor(BLACK);
    if (vatOn) {
      body.appendParagraph('📌 Từ tháng 11/2026: Target (vẫn = TB 3 tháng liền kề × hệ số) và Doanh thu tính % hoàn thành đều là số CHƯA VAT = doanh thu có VAT ÷ 1,08 (VAT 8%). 3 tháng liền kề cũ (có VAT) được quy về chưa VAT trước khi lấy trung bình.')
        .editAsText().setBold(true).setFontSize(10).setForegroundColor(BLACK);
    }
    if (special) {
      body.appendParagraph('⚠ Tháng cao điểm: các Target doanh thu nhân hệ số ' + det.kenhSi.heSoLabel + (det.kenhSi.lyDo ? (' - ' + det.kenhSi.lyDo) : '') + '.')
        .editAsText().setBold(true).setFontSize(10).setForegroundColor(BLACK);
    }
    body.appendParagraph('');

    function h(t) { body.appendParagraph(t).editAsText().setBold(true).setFontSize(12.5).setForegroundColor(BLACK); }
    function sub(t) { body.appendParagraph(t).editAsText().setBold(true).setFontSize(10.5).setForegroundColor(BLACK); }
    function note(t, bold) { body.appendParagraph(t).editAsText().setBold(!!bold).setFontSize(10).setForegroundColor(BLACK); }
    // Bảng dựng bằng appendTable(mảng 2 chiều) -> đúng số dòng/cột, đủ khung, rộng đúng bằng vùng chữ
    function tbl(headers, rows, fracs) {
      const data = [headers].concat(rows.map(function (r) { return r.map(function (x) { return String(x == null ? '' : x); }); }));
      const t = body.appendTable(data);
      t.setBorderWidth(0.75).setBorderColor('#555555');
      for (let c = 0; c < headers.length; c++) t.setColumnWidth(c, Math.round(avail * fracs[c]));
      for (let r = 0; r < data.length; r++) {
        const row = t.getRow(r);
        for (let c = 0; c < headers.length; c++) {
          const cell = row.getCell(c);
          cell.setBackgroundColor(r === 0 ? '#e8e8e8' : (r % 2 === 1 ? '#ffffff' : '#f2f2f2'));   // chữ/lề ô do _kpiFinalizeDoc_ định dạng 1 lần
        }
      }
      body.appendParagraph('');
      return t;
    }

    // ── I. BẢNG TARGET CÁC CHỈ TIÊU ──
    h('I. BẢNG TARGET CÁC CHỈ TIÊU THÁNG ' + mm + '/' + yyyy);
    const kenhSiNote = 'TB3T doanh thu kênh sỉ TOÀN HỆ THỐNG' + vatTag + ' × ' + (det.kenhSi.heSoLabel || '110%');
    const mtNote = 'TB3T doanh số dự án MT' + vatTag + ' × ' + (det.mtGrowth.heSoLabel || '120%');
    tbl(['Vị trí', 'Chỉ tiêu', 'Công thức', 'Target'], [
      ['NVB2B', 'Mục 1 - Phát triển KH mới', 'KH mới = đơn đầu tiên trên 3.000.000đ HOẶC khách đã đặt đến đơn thứ 3. Đạt từ mục tiêu → 900.000đ (+200.000đ/KH vượt); thiếu đúng 1 KH → 450.000đ; còn lại → 0đ', '≥ 2 khách hàng mới đủ điều kiện'],
      ['NVB2B', 'Mục 2 - Doanh số dự án Kênh MT', 'Đạt 100% Target (' + mtNote + ') → 600.000đ; không đạt → 0đ', mtT],
      ['NVB2B', 'Mục 3 - Quỹ phục vụ thành công kênh sỉ', kenhSiNote + '. ≥100% → quỹ 3.000.000đ; 90 đến dưới 100% → 2.400.000đ; 80 đến dưới 90% → 1.500.000đ; dưới 80% → 0đ. Quỹ chia đều cho ' + nB2B + ' NVKD B2B', ksT],
      ['QLB2B', 'Mục 1a - KH mới toàn kênh (tối đa 900.000đ)', 'Tính theo TỔNG KH mới (đủ điều kiện) của các NVKD B2B. Đủ target tổng → 900.000đ; đạt 50% target → 450.000đ; dưới 50% → 0đ', 'Target tổng = 2 KH × ' + nB2B + ' NVKD B2B = ' + (2 * nB2B) + ' KH mới'],
      ['QLB2B', 'Mục 1b - KH mới Quản lý tự kiếm (tối đa 600.000đ)', 'KH mới do chính Quản lý B2B tự kiếm (đủ điều kiện như NVKD B2B): 2 KH → 600.000đ; 1 KH → 200.000đ; không có → 0đ', '≥ 2 khách hàng mới tự kiếm'],
      ['QLB2B', 'Mục 2 - Doanh số phục vụ thành công kênh sỉ (tối đa 1.500.000đ)', 'Dùng chung Target kênh sỉ (không chia quỹ): ≥100% → 1.500.000đ; 90 đến dưới 100% → 1.200.000đ; 80 đến dưới 90% → 750.000đ; dưới 80% → 0đ', ksT],
      ['QLB2B', 'Mục 3 - Tăng trưởng doanh số MT (tối đa 500.000đ)', 'Doanh số kênh MT toàn hệ thống tăng từ 20% trở lên so với bình quân 3 tháng trước (' + mtNote + ') → 500.000đ; không đạt → 0đ', mtT],
      ['QLB2B', 'Mục 4 - Quản trị & tiêu chuẩn vận hành (tối đa 500.000đ)', 'Điểm hành vi TB (thang 100, sau trừ lỗi) của toàn bộ NVKD B2B + CSKH: ≥90% → 500.000đ; 80 đến dưới 90% → 350.000đ; 50 đến dưới 80% → 250.000đ; dưới 50% → 0đ', '≥ 90% điểm hành vi TB'],
      ['CSKH', 'Mục 1 - Quỹ theo Target chung kênh sỉ', 'Dùng chung Target kênh sỉ. ≥100% → quỹ 5.000.000đ; 90 đến dưới 100% → 4.000.000đ; 80 đến dưới 90% → 2.500.000đ; dưới 80% → 0đ. Chia theo hệ số ca (ca sáng = 6, ca tối = 4)', ksT],
      ['CSKH', 'Mục 2 - Tốc độ phản hồi', 'Thời gian xử lý TB (thống kê Fanpage/Zalo): dưới 3 phút → 300.000đ; 3 đến dưới 5 phút → 150.000đ; từ 5 phút → 0đ', 'Dưới 3 phút'],
      ['TPKD', 'Mục 2a - KH mới toàn kênh B2B (tối đa 500.000đ)', 'Tính theo TỔNG KH mới (đủ điều kiện) của các NVKD B2B. Đủ target tổng → 500.000đ; đạt 50% target → 250.000đ; dưới 50% → 0đ', 'Target tổng = 2 KH × ' + nB2B + ' NVKD B2B = ' + (2 * nB2B) + ' KH mới'],
      ['TPKD', 'Mục 2b - Doanh số kênh sỉ toàn hệ thống', '≥100% → 1.000.000đ; 90 đến dưới 100% → 900.000đ; 80 đến dưới 90% → 800.000đ; dưới 80% → 0đ', ksT],
      ['TPKD', 'Mục 2c - Doanh số dự án Kênh MT', 'Đạt 100% Target MT → 500.000đ (nhị phân)', mtT]
    ], [0.12, 0.2, 0.43, 0.25]);
    note('Lưu ý QLB2B: quỹ thưởng tối đa 4.000.000đ/tháng = Mục 1 (1.500.000đ) + Mục 2 (1.500.000đ) + Mục 3 (500.000đ) + Mục 4 (500.000đ).', false);
    note('Lưu ý TPKD: phần B2B tối đa 2.000.000đ. Mục 1 - Kênh B2C (3.000.000đ) tính theo số liệu B2C riêng, KHÔNG thuộc phạm vi văn bản này. Tổng thưởng TPKD tối đa 5.000.000đ.', false);
    body.appendParagraph('');

    // ── II. CHI TIẾT SỐ LIỆU 3 THÁNG ──
    h('II. CHI TIẾT SỐ LIỆU 3 THÁNG DÙNG TÍNH TARGET (minh bạch)');
    function threeMonthTable(title, d, thisVal, thisProj, isCur, pct, thisGross, projGross) {
      sub(title);
      const ms = d.months || [];
      const headers = ['Chỉ tiêu'].concat(ms.map(function (m) { return m.label; }), ['TB ' + (d.monthsUsed || 0) + ' tháng', 'Hệ số', 'Target']);
      const mfr = ms.length ? 0.43 / ms.length : 0.43;
      const fr = [0.16].concat(ms.map(function () { return mfr; }), [0.15, 0.11, 0.15]);
      if (d.vatApplied) {
        // [VAT T11/2026] 2 dòng: có VAT (số gốc) và CHƯA VAT (số dùng tính Target)
        const rowG = ['Doanh thu CÓ VAT'].concat(ms.map(function (m) { return money(m.valueGross); }),
          [d.hasData && d.avgGross != null ? money(d.avgGross) : '-', '-', '-']);
        const rowN = ['Doanh thu CHƯA VAT (tính Target)'].concat(ms.map(function (m) { return money(m.value); }),
          [d.hasData ? money(d.avg) : 'Chưa đủ dữ liệu', d.heSoLabel || '', d.hasData ? money(d.target) : 'Chưa đủ dữ liệu']);
        tbl(headers, [rowG, rowN], fr);
      } else {
        const rowV = ['Giá trị thực tế'].concat(ms.map(function (m) { return money(m.value); }),
          [d.hasData ? money(d.avg) : 'Chưa đủ dữ liệu', d.heSoLabel || '', d.hasData ? money(d.target) : 'Chưa đủ dữ liệu']);
        tbl(headers, [rowV], fr);
      }
      const nt = d.vatApplied ? ' chưa VAT' : '';
      if (d.hasData) note('→ Công thức: TB ' + (d.monthsUsed || 0) + ' tháng' + nt + ' (' + money(d.avg) + ') × ' + d.heSoLabel + ' = Target' + nt + ' (' + money(d.target) + ')', true);
      note('→ Thực hiện lũy kế tháng ' + mm + '/' + yyyy + nt + ': ' + money(thisVal) + (isCur && thisProj != null && !useActualOnly ? ' (Dự kiến cuối tháng: ' + money(thisProj) + ')' : '') + (pct != null ? ' - đạt ' + Number(pct).toFixed(1) + '% Target' : ''), false);
      if (d.vatApplied && thisGross != null) note('→ Đối chiếu: doanh thu CÓ VAT lũy kế = ' + money(thisGross) + (isCur && projGross != null && !useActualOnly ? ' (Dự kiến: ' + money(projGross) + ')' : '') + ' (không dùng để tính %).', false);
      body.appendParagraph('');
    }
    threeMonthTable('Doanh thu kênh sỉ TOÀN HỆ THỐNG (dùng cho NVB2B Mục 3 / QLB2B Mục 2 / CSKH Mục 1 / TPKD Mục 2b):', det.kenhSi, fund.revThis, fund.revProjected, fund.isCurrentMonth, fund.tyLeHoanThanh, fund.revThisGross, fund.revProjectedGross);
    threeMonthTable('Doanh số dự án Kênh MT (dùng cho NVB2B Mục 2 / QLB2B Mục 3 / TPKD Mục 2c):', det.mtGrowth, mt.mtThis, mt.mtProjected, mt.isCurrentMonth, mt.completionPct, mt.mtThisGross, mt.mtProjectedGross);

    // ── III. TIẾN ĐỘ KH MỚI ──
    if (nvs.length) {
      h('III. TIẾN ĐỘ KHÁCH HÀNG MỚI CỦA NVKD B2B');
      tbl(['Nhân viên B2B', 'Mục tiêu KH mới', 'Đã đạt (lũy kế)', 'Dự kiến cuối tháng'],
        nvs.map(function (e) {
          return [e.name, String(Number(e.input.targetNewCust) || 2) + ' KH', String(Math.floor(e.newCustActual || 0)) + ' KH',
            e.isCurrentMonth && !useActualOnly ? (Math.floor(e.newCustProjected || 0) + ' KH') : '-'];
        }), [0.34, 0.22, 0.22, 0.22]);
      const totNew = nvs.reduce(function (a, e) { return a + Math.floor((e.isCurrentMonth && !useActualOnly) ? (e.newCustProjected || 0) : (e.newCustActual || 0)); }, 0);
      note('→ Tổng KH mới toàn kênh B2B: ' + totNew + ' / Target tổng ' + (2 * nB2B) + ' KH (dùng cho QLB2B Mục 1a và TPKD Mục 2a).', true);
    }

    // ── IV. MỐC THƯỞNG ──
    h('IV. BẢNG MỐC THƯỞNG THEO CHÍNH SÁCH KPI KINH DOANH MỚI (tham chiếu chính sách - không đổi theo target)');
    tbl(['Vị trí', 'Mục', 'Mốc thưởng'], [
      ['NVB2B', 'Mục 1 - KH mới', 'Đạt mục tiêu (≥2 KH) → 900.000đ, +200.000đ/KH vượt; thiếu 1 KH → 450.000đ; còn lại → 0đ (KH mới = đơn đầu trên 3.000.000đ hoặc đã đặt đến đơn thứ 3)'],
      ['NVB2B', 'Mục 2 - Dự án MT', 'Đạt 100% Target MT → 600.000đ; không đạt → 0đ'],
      ['NVB2B', 'Mục 3 - Quỹ kênh sỉ', '≥100% → quỹ 3.000.000đ | 90-99% → 2.400.000đ | 80-89% → 1.500.000đ | dưới 80% → 0đ (chia đều cho các NVKD B2B)'],
      ['QLB2B', 'Mục 1a - KH mới toàn kênh', 'Đủ target tổng (2 KH × số NV) → 900.000đ; đạt 50% → 450.000đ; dưới 50% → 0đ'],
      ['QLB2B', 'Mục 1b - KH mới tự kiếm', 'Tự kiếm 2 KH → 600.000đ; 1 KH → 200.000đ'],
      ['QLB2B', 'Mục 2 - Phục vụ kênh sỉ', '≥100% → 1.500.000đ | 90-99% → 1.200.000đ | 80-89% → 750.000đ | dưới 80% → 0đ'],
      ['QLB2B', 'Mục 3 - Tăng trưởng MT', 'Doanh số MT toàn hệ thống tăng từ 20% trở lên so với TB 3 tháng trước → 500.000đ; không đạt → 0đ'],
      ['QLB2B', 'Mục 4 - Quản trị & vận hành', 'Điểm hành vi TB của NVKD B2B + CSKH: ≥90% → 500.000đ | 80-dưới 90% → 350.000đ | 50-dưới 80% → 250.000đ | dưới 50% → 0đ'],
      ['CSKH', 'Mục 1 - Quỹ kênh sỉ (chia theo ca)', '≥100% → quỹ 5.000.000đ | 90-99% → 4.000.000đ | 80-89% → 2.500.000đ | dưới 80% → 0đ. Chia theo hệ số ca: ca sáng = 6, ca tối = 4; 1 hệ số = quỹ kích hoạt ÷ tổng hệ số'],
      ['CSKH', 'Mục 2 - Tốc độ phản hồi', 'TB dưới 3 phút → 300.000đ | 3 đến dưới 5 phút → 150.000đ | từ 5 phút → 0đ'],
      ['TPKD', 'Mục 2a - KH mới toàn kênh B2B', 'Đủ target tổng (2 KH × số NVKD B2B) → 500.000đ | đạt 50% → 250.000đ | dưới 50% → 0đ'],
      ['TPKD', 'Mục 2b - Kênh sỉ toàn hệ thống', '≥100% → 1.000.000đ | 90-99% → 900.000đ | 80-89% → 800.000đ | dưới 80% → 0đ'],
      ['TPKD', 'Mục 2c - Dự án Kênh MT', 'Đạt 100% Target MT → 500.000đ (nhị phân)']
    ], [0.12, 0.2, 0.68]);

    note('Thưởng thực nhận = (Tổng các Mục − Khấu trừ vi phạm) × % Ngày công (% Ngày công = số giờ thực tế ÷ 208 giờ, tối đa 100%).', true);
    note('Khấu trừ vi phạm: 100.000đ/lỗi. Điểm hành vi = 100 − 10 × số lỗi; từ lỗi thứ 5 trở lên (điểm dưới 60/100) → mất toàn bộ thưởng KPI tháng đó.', true);
    note('Doanh thu chỉ được ghi nhận khi đơn đã được Admin/QLB2B duyệt theo quy định hiện hành.', false);
    body.appendParagraph('');

    // ── chữ ký ──
    const sig = body.appendTable([[
      'Trưởng phòng Kinh doanh — Phan Trường Giang\n(Ký, ghi rõ họ tên)',
      'Ban Giám Đốc phê duyệt\n(Ký, ghi rõ họ tên, đóng dấu)'
    ]]);
    sig.setBorderWidth(0.75).setBorderColor('#555555');
    sig.setColumnWidth(0, Math.round(avail / 2)).setColumnWidth(1, Math.round(avail / 2));
    for (let c = 0; c < 2; c++) {
      const cell = sig.getRow(0).getCell(c);
      cell.editAsText().setForegroundColor(BLACK);
      for (let q = 0; q < cell.getNumChildren(); q++) {
        const ch = cell.getChild(q);
        if (ch.getType() === DocumentApp.ElementType.PARAGRAPH) ch.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    }

    if (typeof _kpiFinalizeDoc_ === 'function') _kpiFinalizeDoc_(doc, { keepWidths: true, boldLastCol: false });
    const base64 = _tgtExportDocxB64_(doc);
    return { success: true, base64: base64, fileName: 'ThongBaoTarget_ToanKenhB2B_Thang' + mm + '-' + yyyy + '_9Sach.docx' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    _KPI_FORCE_ACTUAL_ONLY = false;
  }
}