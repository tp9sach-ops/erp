/** ============================================================
 *  9 SẠCH - BỔ SUNG v8.0
 * ============================================================ */

// ── #3: Doanh thu nhân viên so với TB 3 tháng liền kề trước đó ──
function getEmployeeRevenue3Months(monthKey) {
  try {
    function mkey(y, m) { return y + '-' + ('0' + (m + 1)).slice(-2); }
    var parts = (monthKey || '').split('-');
    if (parts.length !== 2) { var n = new Date(); monthKey = mkey(n.getFullYear(), n.getMonth()); parts = monthKey.split('-'); }
    var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;

    var priorKeys = [];
    for (var i = 3; i >= 1; i--) {
      var pm = m - i, py = y;
      while (pm < 0) { pm += 12; py--; }
      priorKeys.push(mkey(py, pm));
    }

    function empRevenueOfMonth(key) {
      var p2 = key.split('-');
      var yy = parseInt(p2[0], 10), mm = parseInt(p2[1], 10) - 1;
      var from = new Date(yy, mm, 1, 0, 0, 0, 0);
      var to = new Date(yy, mm + 1, 0, 23, 59, 59, 999);
      var orders = _readDonHangGrouped(from.toISOString(), to.toISOString(), null, true);
      var byEmp = {};
      orders.forEach(function (o) {
        if (!o.emp || !o.emp.trim()) return;
        byEmp[o.emp] = (byEmp[o.emp] || 0) + o.total;
      });
      return byEmp;
    }

    var curEmp = empRevenueOfMonth(monthKey);
    var priorEmp = priorKeys.map(empRevenueOfMonth);

    var allNames = {};
    Object.keys(curEmp).forEach(function (nm) { allNames[nm] = 1; });
    priorEmp.forEach(function (mp) { Object.keys(mp).forEach(function (nm) { allNames[nm] = 1; }); });

    var result = Object.keys(allNames).map(function (name) {
      var priorVals = priorEmp.map(function (mp) { return mp[name] || 0; });
      var avg = priorVals.reduce(function (a, b) { return a + b; }, 0) / 3;
      var cur = curEmp[name] || 0;
      var pct = avg ? ((cur - avg) / avg * 100) : (cur ? 100 : 0);
      return { name: name, revenue: cur, avgPrior3: avg, pct: pct, priorValues: priorVals };
    }).sort(function (a, b) { return b.revenue - a.revenue; });

    return { monthKey: monthKey, priorKeys: priorKeys, employees: result };
  } catch (e) { return { error: e.toString(), employees: [] }; }
}

// ── #6: Chi tiết đơn theo kênh MT + sửa ngày / xóa ──
function getMTOrdersForChannel(channel, fromDate, toDate) {
  try {
    var sh = _mtStatsSheet();
    var data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    var from = fromDate ? new Date(fromDate) : null;
    var to = toDate ? new Date(toDate) : null;
    if (to) to.setHours(23, 59, 59, 999);
    var out = [];
    data.slice(1).forEach(function (r, i) {
      if (!r[0]) return;
      var d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (from && d < from) return;
      if (to && d > to) return;
      if (channel && (r[1] || '') !== channel) return;
      out.push({
        sheetRow: i + 2, date: _clean(r[0]), channel: r[1] || '', poNo: (r[2] || '').toString(),
        store: _mtStr_(r[3]), boxes: Number(r[4]) || 0, subTotal: Number(r[5]) || 0,
        vat: Number(r[6]) || 0, total: Number(r[7]) || 0, emp: r[9] || '',
        counted: _mtRowCounted(r)
      });
    });
    out.sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
    return out;
  } catch (e) { return []; }
}
function updateMTOrderDate(sheetRow, newDateStr) {
  try {
    var sh = _mtStatsSheet();
    var row = parseInt(sheetRow);
    if (!row || row < 2 || row > sh.getLastRow()) return { success: false, error: 'Dòng không hợp lệ.' };
    const nd = new Date(newDateStr);
    if (isNaN(nd.getTime())) return { success: false, error: 'Ngày không hợp lệ.' };
    sh.getRange(row, 1).setValue(nd);
    try { _kpiBust_(); } catch (e0) {}
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}
function deleteMTOrderRow(sheetRow) {
  try {
    var sh = _mtStatsSheet();
    var row = parseInt(sheetRow);
    if (!row || row < 2 || row > sh.getLastRow()) return { success: false, error: 'Dòng không hợp lệ.' };
    sh.deleteRow(row);
    try { _kpiBust_(); } catch (e0) {}
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ── #7: Doanh thu MT theo tháng, so với TB 3 tháng liền kề ──
function getMTMonthlySeries(monthsCount) {
  try {
    var n = monthsCount || 6;
    var sh = _mtStatsSheet();
    var data = sh.getDataRange().getValues();
    var now = new Date();
    function mkey(y, m) { return y + '-' + ('0' + (m + 1)).slice(-2); }
    var keys = [];
    for (var i = n - 1; i >= 0; i--) {
      var mm = now.getMonth() - i, yy = now.getFullYear();
      while (mm < 0) { mm += 12; yy--; }
      keys.push(mkey(yy, mm));
    }
    var byMonth = {};
    keys.forEach(function (k) { byMonth[k] = 0; });
    if (data.length > 1) {
      data.slice(1).forEach(function (r) {
        if (!r[0]) return;
        var d = r[0] instanceof Date ? r[0] : new Date(r[0]);
        var k = mkey(d.getFullYear(), d.getMonth());
        if (byMonth.hasOwnProperty(k) && _mtRowCounted(r)) byMonth[k] += Number(r[7]) || 0;
      });
    }
    return keys.map(function (k, idx) {
      var priorVals = [];
      for (var j = 1; j <= 3; j++) { if (idx - j >= 0) priorVals.push(byMonth[keys[idx - j]]); }
      var avg = priorVals.length ? (priorVals.reduce(function (a, b) { return a + b; }, 0) / priorVals.length) : 0;
      var pct = avg ? ((byMonth[k] - avg) / avg * 100) : (byMonth[k] ? 100 : 0);
      var p = k.split('-');
      return { month: k, label: p[1] + '/' + p[0], revenue: byMonth[k], avgPrior3: avg, pct: pct };
    });
  } catch (e) { return []; }
}

// ── #8: KPI kết hợp kênh B2B + MT ──
function getCombinedMonthlyKPI(monthKey, channels) {
  try {
    var wantB2B = !channels || channels.indexOf('B2B') !== -1;
    var wantMT = !channels || channels.indexOf('MT') !== -1;
    var b2b = wantB2B ? getMonthlyKPI(monthKey) : null;

    var mtTotal = 0, mtOrders = 0, mtBoxes = 0, mtByChannel = [];
    if (wantMT) {
      var parts = monthKey.split('-');
      var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;
      var from = new Date(y, m, 1, 0, 0, 0, 0), to = new Date(y, m + 1, 0, 23, 59, 59, 999);
      var mtRows = getMTStats(from.toISOString(), to.toISOString());
      var byC = {};
      mtRows.forEach(function (r) {
        var c = (r.channel || '').toUpperCase();
        if (!byC[c]) byC[c] = { channel: c, orders: 0, boxes: 0, total: 0 };
        byC[c].orders++; byC[c].boxes += r.boxes; byC[c].total += r.total;
        mtTotal += r.total; mtOrders++; mtBoxes += r.boxes;
      });
      mtByChannel = Object.keys(byC).map(function (k) { return byC[k]; }).sort(function (a, b) { return b.total - a.total; });
    }
    return { monthKey: monthKey, b2b: b2b, mtTotal: mtTotal, mtOrders: mtOrders, mtBoxes: mtBoxes, mtByChannel: mtByChannel };
  } catch (e) { return { error: e.toString() }; }
}