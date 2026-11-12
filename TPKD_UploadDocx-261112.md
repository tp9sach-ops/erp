/** ============================================================
 *  9 SẠCH - TPKD_UploadDocx.gs  (FILE MỚI - thêm vào project Apps Script)
 *  Đọc file "PhieuDanhGiaKPI_TPKD_YYYY-MM.docx" (phiếu một phần do trang báo cáo B2C xuất)
 *  để lấy số liệu KÊNH B2C còn thiếu cho KPI Trưởng phòng Kinh doanh:
 *    - Doanh số B2C thực hiện (Mục 1a/1b)
 *    - Target hòa vốn toàn hệ thống (Mục 1a)
 *    - TB lũy kế 3 tháng liền kề (Mục 1b)
 *  Nguồn: bảng trong phiếu + "Mã chuyển tiếp" KPITPKD1:... (kiểm tra tháng, % đạt để đối chiếu).
 *  Chỉ trả số liệu về trang; việc tính thưởng vẫn do getTPKDData() làm như cũ.
 * ============================================================ */
function parseTPKDPartialDoc(pw, base64) {
  try {
    if (String(pw || '').trim() !== '9999') return { success: false, error: 'Không có quyền.' };
    if (!base64) return { success: false, error: 'File rỗng.' };
    var files = Utilities.unzip(Utilities.newBlob(Utilities.base64Decode(base64), 'application/zip', 'phieu.docx'));
    var xml = '';
    for (var i = 0; i < files.length; i++) {
      if (files[i].getName() === 'word/document.xml') { xml = files[i].getDataAsString('UTF-8'); break; }
    }
    if (!xml) return { success: false, error: 'Không đọc được nội dung file .docx.' };

    var toText = function (x) {
      return x.replace(/<\/w:p>/g, '\n').replace(/<w:br\s*\/>/g, '\n').replace(/<w:tab\s*\/>/g, '\t')
        .replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
    };
    var vnd = function (s) { var d = String(s || '').replace(/\D/g, ''); return d ? parseInt(d, 10) : 0; };

    // 1) Mã chuyển tiếp
    var full = toText(xml), meta = null;
    var m = full.match(/KPITPKD1:([A-Za-z0-9+\/=]+)/);
    if (m) {
      try { meta = JSON.parse(Utilities.newBlob(Utilities.base64Decode(m[1])).getDataAsString('UTF-8')); } catch (e) { meta = null; }
    }
    if (meta && meta.loai && meta.loai !== 'KPI_TPKD') return { success: false, error: 'Đây không phải phiếu KPI Trưởng phòng Kinh doanh.' };

    // 2) Bảng chi tiết
    var rows = xml.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
    var r1a = null, r1b = null;
    rows.forEach(function (r) {
      var cells = (r.match(/<w:tc[ >][\s\S]*?<\/w:tc>/g) || []).map(function (c) { return toText(c).trim(); });
      if (!cells.length) return;
      if (/^Mục\s*1a/i.test(cells[0])) r1a = cells;
      if (/^Mục\s*1b/i.test(cells[0])) r1b = cells;
    });
    if (!r1a && !r1b) return { success: false, error: 'Không tìm thấy Mục 1a / 1b trong file. Hãy dùng đúng file PhieuDanhGiaKPI_TPKD do trang báo cáo B2C xuất.' };

    var be = 0, avg = 0, b2c = 0, hs = 100;
    if (r1a) {
      var mb = (r1a[1] || '').match(/Target[^:]*:\s*([\d.,]+)/i); if (mb) be = vnd(mb[1]);
      b2c = vnd(r1a[2]);
    }
    if (r1b) {
      var mt = (r1b[1] || '').match(/TB3T\s*([\d.,]+)\s*đ?\s*[×x]\s*(\d+(?:[.,]\d+)?)\s*%/i);
      if (mt) { avg = vnd(mt[1]); hs = parseFloat(String(mt[2]).replace(',', '.')) || 100; }
      else { var mt2 = (r1b[1] || '').match(/TB3T\s*([\d.,]+)/i); if (mt2) avg = vnd(mt2[1]); }
      if (!b2c) b2c = vnd(r1b[2]);
    }
    var month = (meta && meta.thang) ? String(meta.thang) : '';
    if (!month) { var mm = full.match(/THÁNG\s*(\d{1,2})\/(\d{4})/i); if (mm) month = mm[2] + '-' + ('0' + mm[1]).slice(-2); }

    return {
      success: true, month: month, b2c: b2c, be: be, avgB2C: avg,
      heSoThang: (meta && meta.heSoThang) ? meta.heSoThang : hs,
      pct1a: meta && meta.b2c && meta.b2c['1a'] ? meta.b2c['1a'].dat : null,
      pct1b: meta && meta.b2c && meta.b2c['1b'] ? meta.b2c['1b'].dat : null,
      hasMeta: !!meta
    };
  } catch (e) { return { success: false, error: e.toString() }; }
}