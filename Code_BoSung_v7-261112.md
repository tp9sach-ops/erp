/** ============================================================
 *  9 SẠCH - BỔ SUNG v7.0
 *  ================================================================
 *  CÁCH DÙNG: Trong Apps Script Editor, bấm dấu "+" cạnh "Files" ->
 *  Script -> đặt tên "Code_BoSung_v7" -> dán TOÀN BỘ nội dung file
 *  này vào. KHÔNG cần sửa gì trong Code.gs hiện tại - các hàm dùng
 *  lại (_readDonHangGrouped, _loadTongHopMap, _monthKey, _clean...)
 *  đã có sẵn trong Code.gs.
 *
 *  NỘI DUNG:
 *  1) getMonthlyKPI(monthKey) - KPI doanh thu cho 1 tháng bất kỳ do
 *     người dùng chọn (dùng cho bộ lọc tháng ở trang Sức Khỏe KD).
 *  2) getCustomerOrdersDetailed(phoneOrName, monthKey) - danh sách
 *     đơn hàng đầy đủ (ngày, SĐT, tên, nhân viên, tổng tiền, vị trí
 *     dòng để xóa) của 1 khách hàng trong 1 tháng (dùng cho trang
 *     Khách Hàng - tìm kiếm khách + xem đơn theo tháng).
 * ============================================================ */

// ════════════════════════════════════════════════════════════
// 1) KPI DOANH THU THEO THÁNG BẤT KỲ (chọn tháng ở trang BI)
// ════════════════════════════════════════════════════════════
function getMonthlyKPI(monthKey) {
  try {
    function mkey(yr, mo) { return yr + '-' + ('0' + (mo + 1)).slice(-2); }

    var parts = (monthKey || '').split('-');
    if (parts.length !== 2) monthKey = mkey(new Date().getFullYear(), new Date().getMonth());
    parts = monthKey.split('-');
    var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;

    var lmY = y, lmM = m - 1;
    if (lmM < 0) { lmM = 11; lmY--; }

    var tongHop = _loadTongHopMap();
    function statFromMap(key) {
      return tongHop[key] ? Object.assign({}, tongHop[key]) : { revenue: 0, orders: 0, qty: 0, aov: 0, customers: 0 };
    }
    function pct(cur, prev) { if (!prev) return cur ? 100 : 0; return (cur - prev) / prev * 100; }

    var thisMonth = statFromMap(mkey(y, m));
    var lastMonth = statFromMap(mkey(lmY, lmM));

    var from = new Date(y, m, 1, 0, 0, 0, 0);
    var to = new Date(y, m + 1, 0, 23, 59, 59, 999);
    var orders = _readDonHangGrouped(from.toISOString(), to.toISOString(), null, true);

    var byEmp = {}, byPriceType = {}, byProduct = {};
    var rev = 0, qty = 0; var cust = {};
    orders.forEach(function (o) {
      rev += o.total;
      if (o.phone) cust[o.phone] = 1;
      if (!byEmp[o.emp]) byEmp[o.emp] = { orders: 0, revenue: 0 };
      byEmp[o.emp].orders++; byEmp[o.emp].revenue += o.total;
      if (!byPriceType[o.priceType]) byPriceType[o.priceType] = { orders: 0, revenue: 0 };
      byPriceType[o.priceType].orders++; byPriceType[o.priceType].revenue += o.total;
      o.items.forEach(function (it) {
        qty += it.qty;
        if (!byProduct[it.name]) byProduct[it.name] = { qty: 0, revenue: 0 };
        byProduct[it.name].qty += it.qty; byProduct[it.name].revenue += it.price;
      });
    });

    // Nếu TongHopThang chưa có dòng của tháng này (chưa chạy đồng bộ),
    // tự tính bù trực tiếp từ DonHang để KPI vẫn hiển thị đúng.
    if (thisMonth.revenue === 0 && orders.length > 0) {
      thisMonth = {
        revenue: rev, orders: orders.length, qty: qty,
        aov: orders.length ? Math.round(rev / orders.length) : 0,
        customers: Object.keys(cust).length
      };
    }

    var empArr = Object.keys(byEmp).map(function (k) {
      return { name: k, orders: byEmp[k].orders, revenue: byEmp[k].revenue };
    }).sort(function (a, b) { return b.revenue - a.revenue; });

    var priceTypeArr = Object.keys(byPriceType).map(function (k) {
      return { name: k, orders: byPriceType[k].orders, revenue: byPriceType[k].revenue };
    }).sort(function (a, b) { return b.revenue - a.revenue; });

    var productArr = Object.keys(byProduct).map(function (k) {
      return { name: k, qty: byProduct[k].qty, revenue: byProduct[k].revenue };
    }).sort(function (a, b) { return b.revenue - a.revenue; }).slice(0, 10);

    return {
      monthKey: monthKey, thisMonth: thisMonth, lastMonth: lastMonth,
      momPct: pct(thisMonth.revenue, lastMonth.revenue),
      empArr: empArr, priceTypeArr: priceTypeArr, productArr: productArr
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ════════════════════════════════════════════════════════════
// 2) CHI TIẾT ĐƠN HÀNG CỦA 1 KHÁCH TRONG 1 THÁNG (trang Khách Hàng)
//    Trả về đủ: ngày, SĐT, tên, nhân viên, tổng tiền, và vị trí dòng
//    (sheetStartRow/rowCount) để có thể xóa (hủy) đơn trực tiếp.
// ════════════════════════════════════════════════════════════
function getCustomerOrdersDetailed(phoneOrName, monthKey) {
  try {
    var parts = (monthKey || '').split('-');
    var from = null, to = null;
    if (parts.length === 2) {
      var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;
      from = new Date(y, m, 1, 0, 0, 0, 0);
      to = new Date(y, m + 1, 0, 23, 59, 59, 999);
    }
    var orders = _readDonHangGrouped(from ? from.toISOString() : null, to ? to.toISOString() : null, null);
    var key = (phoneOrName || '').toString().trim();
    var rows = orders.filter(function (o) {
      return (o.phone || '').toString() === key || (o.name || '').toString() === key;
    }).map(function (o) {
      return {
        date: o.date, phone: o.phone, name: o.name, emp: o.emp,
        qty: o.items.reduce(function (a, it) { return a + it.qty; }, 0),
        total: o.total, sheetStartRow: o.sheetStartRow, rowCount: o.items.length
      };
    });
    rows.sort(function (a, b) { return new Date(a.date) - new Date(b.date); });
    return rows;
  } catch (e) { return []; }
}
// [MỚI] Lấy đơn hàng của 1 khách theo khoảng ngày cụ thể (thay vì theo
// tháng) - dùng cho khung tìm kiếm khách hàng lọc theo Từ ngày/Đến ngày.
// Nếu fromDate/toDate rỗng -> trả về TOÀN BỘ đơn hàng của khách đó.
function getCustomerOrdersInRange(phoneOrName, fromDate, toDate) {
  try {
    var orders = _readDonHangGrouped(fromDate || null, toDate || null, null);
    var key = (phoneOrName || '').toString().trim();
    var rows = orders.filter(function (o) {
      return (o.phone || '').toString() === key || (o.name || '').toString() === key;
    }).map(function (o) {
      return {
        date: o.date, phone: o.phone, name: o.name, emp: o.emp,
        qty: o.items.reduce(function (a, it) { return a + it.qty; }, 0),
        total: o.total, sheetStartRow: o.sheetStartRow, rowCount: o.items.length
      };
    });
    rows.sort(function (a, b) { return new Date(a.date) - new Date(b.date); });
    return rows;
  } catch (e) { return []; }
}