/** ============================================================
 *  9 SẠCH - XUẤT PHIẾU ĐÁNH GIÁ KPI CÁ NHÂN (Prompt 4)  [đã gắn _kpiFinalizeDoc_]
 *  Dùng lại đúng logic tính đã có ở getKPIBonusList() / 
 *  getKPIBonusListCSKH() - không tính lại công thức ở nơi khác.
 * ============================================================ */
function exportKPIEvaluationDoc(monthKey, empName, role, useActualOnly){
  _KPI_FORCE_ACTUAL_ONLY = !!useActualOnly;
  try{
    var isCSKH = role === 'CSKH';
    var entry = null;
    if(isCSKH){
      var listC = getKPIBonusListCSKH(monthKey);
      if(listC && listC.error) return {success:false, error:listC.error};
      entry = (listC||[]).find(function(e){ return e.name===empName; });
    } else {
      var listB = getKPIBonusList(monthKey);
      if(listB && listB.error) return {success:false, error:listB.error};
      entry = (listB||[]).find(function(e){ return e.name===empName; });
    }
    if(!entry) return {success:false, error:'Không tìm thấy dữ liệu KPI của "'+empName+'" trong tháng '+monthKey+'.'};
    if(entry.error) return {success:false, error:'Dữ liệu KPI của "'+empName+'" bị lỗi khi tính toán: '+entry.error};

    if(useActualOnly){
      if(entry.fund) entry.fund.isCurrentMonth=false;
      if(entry.mtGrowth) entry.mtGrowth.isCurrentMonth=false;
      if(entry.kenhSiFund) entry.kenhSiFund.isCurrentMonth=false;
      entry.isCurrentMonth=false;
    }

    var parts = monthKey.split('-');
    var yyyy = parts[0], mm = parts[1];
    var now = new Date();
    var ngayStr = 'ngày ' + now.getDate() + ' tháng ' + (now.getMonth()+1) + ' năm ' + now.getFullYear();
    var roleLabel = isCSKH ? 'Chăm Sóc Khách Hàng' : (entry.role==='QLB2B' ? 'Quản Lý B2B' : 'Nhân Viên B2B');
    var money = function(v){ return Math.round(v||0).toLocaleString('vi-VN')+'đ'; };
    // [VAT T11/2026] từ T11/2026: Target + Doanh thu tính % đều CHƯA VAT; phiếu hiển thị thêm cột doanh thu CÓ VAT để đối chiếu
    var vatOn = _kpiVatOn_(monthKey);
    var vatTag = vatOn ? ' (chưa VAT)' : '';
    var revTxt = function(o, kThis, kProj, kThisG, kProjG){
      var cur = o && o.isCurrentMonth;
      if(!vatOn) return money(o[kThis]) + (cur ? (' (Dự kiến: '+money(o[kProj])+')') : '');
      return 'Chưa VAT: '+money(o[kThis]) + (cur ? (' (Dự kiến: '+money(o[kProj])+')') : '')
        + '\nCó VAT: '+money(o[kThisG]) + (cur ? (' (Dự kiến: '+money(o[kProjG])+')') : '');
    };

    var doc = DocumentApp.create('TEMP_PHIEUKPI_' + empName.replace(/[^\w]/g,'') + '_' + monthKey + '_' + new Date().getTime());
    var body = doc.getBody();
    body.setMarginTop(45).setMarginBottom(45).setMarginLeft(55).setMarginRight(55);

    _appendKPIDocHeader(body, 'KPI-'+mm+yyyy, ngayStr);

    var pTitle = body.appendParagraph('PHIẾU ĐÁNH GIÁ KPI THÁNG '+mm+'/'+yyyy);
    pTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pTitle.editAsText().setBold(true).setFontSize(16).setForegroundColor('#000000');
    var pSub = body.appendParagraph(empName+' — '+roleLabel);
    pSub.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    pSub.editAsText().setBold(true).setFontSize(12).setForegroundColor('#000000');
    body.appendParagraph('');

    var pMode = body.appendParagraph(useActualOnly
      ? '📌 Tính theo LŨY KẾ THỰC TẾ đến thời điểm xuất (không có phần dự kiến).'
      : '📌 Có bao gồm phần DỰ KIẾN hoàn thành đến cuối tháng cho các chỉ tiêu tháng hiện tại.');
    pMode.editAsText().setBold(true).setFontSize(10).setForegroundColor('#000000');
    if(vatOn){
      body.appendParagraph('📌 Từ tháng 11/2026: Target và Doanh thu dùng tính % hoàn thành đều là số CHƯA VAT (doanh thu có VAT ÷ 1,08; VAT 8%). Bảng có thêm số doanh thu CÓ VAT để đối chiếu.')
        .editAsText().setBold(true).setFontSize(10).setForegroundColor('#000000');
    }
    body.appendParagraph('');

    body.appendParagraph('I. CHI TIẾT CÁC MỤC KPI').editAsText().setBold(true).setFontSize(13).setForegroundColor('#000000');
    var table = body.appendTable();
    var h = table.appendTableRow();
    ['Mục','Hình thức tính / Target','Thực hiện','% Hoàn thành','Thưởng'].forEach(function(t){
      var c=h.appendTableCell(t);
      c.editAsText().setBold(true).setForegroundColor('#000000').setFontSize(9.5);
      c.setBackgroundColor('#e8e8e8');
    });
    var _rowIdx=0;
    function addRow(name, formula, actual, pct, bonus){
      var r=table.appendTableRow();
      var bg = (_rowIdx++%2===0) ? '#ffffff' : '#f2f2f2';
      [name, formula, actual, pct].forEach(function(val){
        var c=r.appendTableCell(val); c.setBackgroundColor(bg);
      });
      var bonusCell=r.appendTableCell(bonus);
      bonusCell.setBackgroundColor(bg);   // font/đậm do _kpiFinalizeDoc_ định dạng 1 lần
    }

    var totalRawVal, totalDeductVal, totalFinalVal, blocked, heSoNgayCong, opsPart = 0;

    if(isCSKH){
      var kf=entry.kenhSiFund||{}, sh2=entry.shift||{}, sp=entry.speed;
      var caLbl = entry.ca==='Toi' ? 'Ca tối (hệ số 4)' : 'Ca sáng (hệ số 6)';
      addRow('Mục 1 - Quỹ theo Target chung kênh sỉ',
        'Target chung = TB3T doanh thu kênh sỉ'+vatTag+' × hệ số. Quỹ kích hoạt '+money(kf.quyCSKH)+' chia theo hệ số ca (tổng hệ số cả đội '+(sh2.totalHeSo||0)+') — '+caLbl+'. Target'+vatTag+': '+(kf.target?money(kf.target):'Chưa đủ dữ liệu'),
        revTxt(kf,'revThis','revProjected','revThisGross','revProjectedGross'),
        (kf.tyLeHoanThanh||0).toFixed(1)+'%',
        money(entry.calc.salesBonus));
      addRow('Mục 2 - Tốc độ phản hồi',
        'Thời gian xử lý TB (thống kê Fanpage/Zalo): dưới 3 phút → 300.000đ; 3 đến dưới 5 phút → 150.000đ; từ 5 phút → 0đ',
        sp ? ('TB '+sp.avgText) : 'Chưa có dữ liệu (chưa upload file thống kê)',
        entry.calc.speedTier || '-',
        money(entry.calc.speedBonus));
      totalRawVal = entry.calc.rawTotal; totalDeductVal = entry.calc.deduction; totalFinalVal = entry.calc.finalTotal; blocked = entry.calc.blocked;
      heSoNgayCong = entry.calc.heSoNgayCong;

      body.appendParagraph('');
      body.appendParagraph('II. KHẤU TRỪ VI PHẠM (100.000đ/lỗi)').editAsText().setBold(true).setFontSize(12).setForegroundColor('#000000');
      var vTable=body.appendTable();
      var vh=vTable.appendTableRow();
      ['Số lỗi vi phạm','Số tiền trừ/lỗi','Tổng tiền trừ'].forEach(function(t){
        var c=vh.appendTableCell(t); c.editAsText().setBold(true).setForegroundColor('#000000'); c.setBackgroundColor('#e8e8e8');
      });
      var vr=vTable.appendTableRow();
      vr.appendTableCell(String(entry.input.cskhViolations||0)).editAsText().setForegroundColor('#000000');
      vr.appendTableCell('100.000đ').editAsText().setForegroundColor('#000000');
      vr.appendTableCell(money(entry.calc.deduction)).editAsText().setForegroundColor('#000000');
    } else {
      var fund=entry.fund||{}, mtG=entry.mtGrowth||{};
      var isManager = entry.role==='QLB2B';
      if(isManager){
        addRow('Mục 1 - Phát triển KH mới toàn kênh',
          'Tổng KH mới (đủ điều kiện) của các NVKD B2B. Target tổng = 2 KH × '+(entry.totalB2BCount||1)+' NV = '+(entry.calc.teamTarget||0)+' KH. Đủ → 900.000đ; đạt 50% → 450.000đ; dưới 50% → 0đ',
          String(entry.newCustActual||0)+' KH'+(entry.isCurrentMonth?(' (Dự kiến: '+Math.round(entry.newCustProjected||0)+')'):''),
          '-', money(entry.calc.newCustBonus));
        addRow('Mục 2 - Doanh số phục vụ thành công kênh sỉ',
          'Dùng chung Target kênh sỉ, không chia quỹ. Target'+vatTag+': '+(fund.target?money(fund.target):'Chưa đủ dữ liệu'),
          revTxt(fund,'revThis','revProjected','revThisGross','revProjectedGross'),
          (fund.tyLeHoanThanh||0).toFixed(1)+'%', money(entry.calc.serviceBonus));
      } else {
        addRow('Mục 1 - Phát triển khách hàng mới',
          'KH mới = đơn đầu tiên >3.000.000đ hoặc đã đặt đến đơn thứ 3. Mốc: ≥'+(entry.input.targetNewCust||2)+' KH → 900.000đ (+200.000đ/KH vượt); thiếu 1 KH → 450.000đ',
          String(entry.newCustActual||0)+' KH'+(entry.isCurrentMonth?(' (Dự kiến: '+Math.round(entry.newCustProjected||0)+')'):''),
          '-', money(entry.calc.newCustBonus));
      }
      addRow(isManager?'Mục 3 - Tăng trưởng doanh số dự án MT':'Mục 2 - Doanh số dự án Kênh MT',
        'Nhị phân: tăng ≥20% so với TB3T. Target = TB3T × hệ số tăng trưởng'+vatTag+': '+(mtG.mtTargetAvg3?money(mtG.mtTargetAvg3):'Chưa đủ dữ liệu'),
        money(entry.input.projectCompleted)+((vatOn && mtG.mtThisGross!=null) ? ('\nChưa VAT (tính %) / Có VAT: '+money(mtG.mtThis)+' / '+money(mtG.mtThisGross)) : ''), entry.calc.projectPct.toFixed(1)+'%',
        money(entry.calc.projectBonus));
      if(!isManager){
        addRow('Mục 3 - Quỹ phục vụ thành công kênh sỉ',
          'TB3T doanh thu kênh sỉ toàn hệ thống'+vatTag+' × hệ số, chia đều cho NVKD B2B. Target'+vatTag+': '+(fund.target?money(fund.target):'Chưa đủ dữ liệu'),
          revTxt(fund,'revThis','revProjected','revThisGross','revProjectedGross'),
          (fund.tyLeHoanThanh||0).toFixed(1)+'%', money(entry.calc.kenhSiBonus));
      } else {
        addRow('Mục 4 - Quản trị & tiêu chuẩn vận hành',
          'Điểm hành vi TB (sau trừ lỗi) của toàn bộ NVKD B2B + CSKH: ≥80% → 1.000.000đ; 50% đến dưới 80% → 500.000đ; dưới 50% → 0đ',
          (entry.calc.teamPct||0).toFixed(1)+'%', '-', money(entry.calc.opsBonus));
      }
      totalRawVal = entry.calc.rawBonus;
      totalDeductVal = entry.calc.deduction; totalFinalVal = entry.calc.grandTotal; blocked = entry.calc.blocked;
      heSoNgayCong = entry.calc.heSoNgayCong;

      body.appendParagraph('');
      body.appendParagraph('II. KHẤU TRỪ VI PHẠM (100.000đ/lỗi)').editAsText().setBold(true).setFontSize(12).setForegroundColor('#000000');
      var vTable2=body.appendTable();
      var vh2=vTable2.appendTableRow();
      ['Loại lỗi','Số lượng','Số tiền trừ'].forEach(function(t){
        var c=vh2.appendTableCell(t); c.editAsText().setBold(true).setForegroundColor('#000000'); c.setBackgroundColor('#e8e8e8');
      });
      [['Không chủ động kế hoạch',entry.input.viol1],['Không nhắc đơn đại lý',entry.input.viol2],
       ['Xử lý chậm >3 giờ',entry.input.viol3],['Công nợ trễ hạn',entry.input.viol4]].forEach(function(v){
        var r=vTable2.appendTableRow();
        r.appendTableCell(v[0]).editAsText().setForegroundColor('#000000');
        r.appendTableCell(String(v[1]||0)).editAsText().setForegroundColor('#000000');
        r.appendTableCell(money(100000*(v[1]||0))).editAsText().setForegroundColor('#000000');
      });
      var totalViolRow=vTable2.appendTableRow();
      totalViolRow.appendTableCell('Tổng số lỗi').editAsText().setBold(true).setForegroundColor('#000000');
      totalViolRow.appendTableCell(String(entry.calc.totalViolations||0)).editAsText().setBold(true).setForegroundColor('#000000');
      totalViolRow.appendTableCell(money(entry.calc.deduction)).editAsText().setBold(true).setForegroundColor('#000000');
    }

    if(blocked){
      body.appendParagraph('');
      var pB=body.appendParagraph('⚠ MẤT TOÀN BỘ THƯỞNG KPI THÁNG NÀY DO VI PHẠM QUÁ NGƯỠNG (dưới 60/100 điểm).');
      pB.editAsText().setBold(true).setForegroundColor('#000000');
    }

    body.appendParagraph('');
    body.appendParagraph('III. TỔNG KẾT THƯỞNG').editAsText().setBold(true).setFontSize(13).setForegroundColor('#000000');
    var sumTable = body.appendTable();
    function sumRow(label, valText, highlight){
      var r=sumTable.appendTableRow();
      var lc=r.appendTableCell(label); lc.editAsText().setBold(true).setFontSize(highlight?12:10).setForegroundColor('#000000');
      var vc=r.appendTableCell(valText);
      vc.editAsText().setBold(true).setFontSize(highlight?13:10).setForegroundColor('#000000');
      if(highlight){ lc.setBackgroundColor('#e8e8e8'); vc.setBackgroundColor('#e8e8e8'); }
    }
    sumRow('Tổng các mục (trước khấu trừ)', money(totalRawVal), false);
    sumRow('Tổng khấu trừ vi phạm', '-'+money(totalDeductVal), false);
    sumRow('% Ngày công ('+(entry.input.gioCongThucTe||208)+' giờ ÷ 208 giờ)', ((heSoNgayCong||1)*100).toFixed(1)+'%', false);
    sumRow('TỔNG THỰC NHẬN = (Tổng các mục − Khấu trừ) × % Ngày công', money(totalFinalVal), true);

    if(entry.input.note){
      body.appendParagraph('');
      body.appendParagraph('Ghi chú: '+entry.input.note).editAsText().setForegroundColor('#000000');
    }

    // [MỚI] Chữ ký: Nhân viên - Trưởng phòng Kinh doanh - Ban Giám Đốc
    body.appendParagraph('');
    body.appendParagraph('');
    var sigTable = body.appendTable();
    var sigRow = sigTable.appendTableRow();
    var c1=sigRow.appendTableCell('Nhân viên nhận thưởng — '+empName+'\n(Ký, ghi rõ họ tên)\n\n\n\n');
    var c2=sigRow.appendTableCell('Trưởng phòng Kinh doanh — Phan Trường Giang\n(Ký, ghi rõ họ tên)\n\n\n\n');
    var c3=sigRow.appendTableCell('Ban Giám Đốc\n(Ký, ghi rõ họ tên, đóng dấu)\n\n\n\n');
    [c1,c2,c3].forEach(function(c){
      c.editAsText().setForegroundColor('#000000');
      c.setWidth(165);
      _kpiCenterCell_(c);
    });

    _kpiFinalizeDoc_(doc);   // [MỚI] chuẩn hóa font/đậm/nghiêng + độ rộng cột để in rõ (hàm ở file KPI_DocStyle.gs)
    var base64 = _exportDocAsBase64Docx(doc);
    return { success:true, base64: base64, fileName: 'PhieuDanhGiaKPI_'+empName.replace(/[^\w]/g,'')+'_Thang'+mm+'-'+yyyy+'.docx' };
  } catch(e){ return { success:false, error:e.toString() }; }
}