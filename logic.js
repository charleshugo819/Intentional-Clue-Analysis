/* ============================================================
 * logic.js — 线索数据处理核心逻辑（UMD：浏览器 + Node 通用）
 * 依赖全局 XLSX（SheetJS）。在浏览器由 CDN 提供；在 Node 中需先
 *   global.XLSX = require('xlsx');
 * 对外暴露：LeadsProcessor.processWorkbook(wb)
 * ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.LeadsProcessor = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var XHS = '小红书';
  var SSY = '水师营';
  var SX = '线上表格';
  var SSY_SX = SSY + '/' + SX;

  function sheetToRows(ws) {
    if (!ws || !ws['!ref']) return [];
    var range = XLSX.utils.decode_range(ws['!ref']);
    var out = [];
    for (var r = range.s.r; r <= range.e.r; r++) {
      var row = [];
      for (var c = range.s.c; c <= range.e.c; c++) {
        var addr = XLSX.utils.encode_cell({ r: r, c: c });
        var cell = ws[addr];
        row.push(cell ? cell.v : undefined);
      }
      out.push(row);
    }
    return out;
  }

  function isRealPhone(v) {
    if (v === null || v === undefined) return false;
    if (typeof v === 'number') return true;
    var s = String(v).trim();
    return s !== '' && /^\d+$/.test(s);
  }

  function mask(v) {
    if (v === null || v === undefined) return '****';
    var s = String(v).trim();
    if (s === '') return '****';
    return s.slice(0, 3) + '****' + s.slice(7);
  }

  function findCol(headerRow, name) {
    for (var c = 0; c < headerRow.length; c++) {
      var h = headerRow[c];
      if (h !== undefined && h !== null && String(h).indexOf(name) !== -1) return c;
    }
    return -1;
  }

  function uniqueNames(names) {
    var seen = {}, out = [];
    (names || []).forEach(function (n) {
      if (!seen[n]) { seen[n] = true; out.push(n); }
    });
    return out;
  }

  function processWorkbook(wb) {
    var crm = wb.Sheets['CRM'];
    var ssy = wb.Sheets['水师营'];
    var sx = wb.Sheets['线上表'];
    var pn = wb.Sheets['人员名单'];
    if (!crm || !ssy || !sx || !pn) {
      throw new Error('文件缺少必要工作表：CRM / 水师营 / 线上表 / 人员名单');
    }

    /* ---------- CRM（数据自第3行起） ---------- */
    var cr = sheetToRows(crm);
    var crMobile = findCol(cr[0] || [], '手机号码密文');
    var crAct = findCol(cr[0] || [], '活动名称');
    var crTime = findCol(cr[0] || [], '创建时间');
    var crmMobile = {};
    for (var i = 2; i < cr.length; i++) {
      var row = cr[i];
      var m = row[crMobile];
      if (m === undefined || m === null || String(m).trim() === '') continue;
      var key = String(m).trim();
      if (!crmMobile[key]) crmMobile[key] = [];
      crmMobile[key].push({ act: crAct >= 0 ? row[crAct] : undefined, time: crTime >= 0 ? row[crTime] : undefined });
    }
    function crmAct(phone) {
      var arr = crmMobile[phone];
      if (!arr) return null;
      for (var a = 0; a < arr.length; a++) {
        if (arr[a].act !== undefined && arr[a].act !== null && String(arr[a].act).trim() !== '') return arr[a].act;
      }
      return null;
    }

    /* ---------- 水师营（数据自第2行起） ---------- */
    var sy = sheetToRows(ssy);
    var syMobile = findCol(sy[0] || [], '学员手机');
    var syName = findCol(sy[0] || [], '录入人姓名');
    var ssyByPhone = {};
    for (i = 1; i < sy.length; i++) {
      row = sy[i];
      m = row[syMobile];
      if (m === undefined || m === null || String(m).trim() === '') continue;
      key = String(m).trim();
      if (!ssyByPhone[key]) ssyByPhone[key] = [];
      var nm = syName >= 0 ? row[syName] : undefined;
      if (nm !== undefined && nm !== null && String(nm).trim() !== '') ssyByPhone[key].push(String(nm).trim());
    }

    /* ---------- 线上表（数据自第2行起） ---------- */
    var xr = sheetToRows(sx);
    var xNum = findCol(xr[0] || [], '线索号码');
    var xOp = findCol(xr[0] || [], '账号运营人');
    var xPlat = findCol(xr[0] || [], '来源平台');
    var sxList = [];
    for (i = 1; i < xr.length; i++) {
      row = xr[i];
      sxList.push({
        num: xNum >= 0 ? row[xNum] : undefined,
        op: xOp >= 0 ? row[xOp] : undefined,
        plat: xPlat >= 0 ? row[xPlat] : undefined
      });
    }

    /* ---------- 人员名单（标题行1，表头行2，数据行3起） ---------- */
    var pr = sheetToRows(pn);
    var pCity = findCol(pr[1] || [], '地市');
    var pName = findCol(pr[1] || [], '姓名');
    var pRole = findCol(pr[1] || [], '新分岗分类');
    var pTarget = findCol(pr[1] || [], '意向目标');
    var pnInput = [];
    for (i = 2; i < pr.length; i++) {
      row = pr[i];
      var pn = pName >= 0 ? row[pName] : undefined;
      if (pn === undefined || pn === null || String(pn).trim() === '') continue;
      pnInput.push({
        city: pCity >= 0 ? row[pCity] : undefined,
        name: String(pn).trim(),
        role: pRole >= 0 ? row[pRole] : undefined,
        target: pTarget >= 0 ? row[pTarget] : undefined
      });
    }

    /* ---------- 结果1 = 纯手机号匹配（不做时间过滤） ---------- */
    var R1 = Object.keys(ssyByPhone).filter(function (m) { return crmMobile.hasOwnProperty(m); });
    var R1set = {};
    R1.forEach(function (m) { R1set[m] = true; });
    var sxRealMasked = {};
    sxList.forEach(function (s) { if (isRealPhone(s.num)) sxRealMasked[mask(s.num)] = true; });

    /* ---------- 结果2 = 并集，交集去重一次 ---------- */
    var entries = [];
    R1.forEach(function (m) {
      if (sxRealMasked.hasOwnProperty(m)) return;
      var names = uniqueNames(ssyByPhone[m]);
      if (names.length === 0) names = [null];
      names.forEach(function (n) {
        entries.push({ f: m, g: n, j: crmAct(m), k: XHS, src: 'water_only' });
      });
    });
    sxList.forEach(function (s) {
      var f = mask(s.num);
      var k = (s.plat !== undefined && s.plat !== null && String(s.plat).trim() !== '') ? String(s.plat).trim() : XHS;
      var isReal = isRealPhone(s.num);
      var j = (isReal && R1set.hasOwnProperty(f)) ? crmAct(f) : null;
      var src = R1set.hasOwnProperty(f) ? 'inter' : 'sx_only';
      entries.push({ f: f, g: s.op, j: j, k: k, src: src });
    });

    /* ---------- 完成量 / 完成率 ---------- */
    var nameDone = {};
    entries.forEach(function (e) {
      if (e.g !== undefined && e.g !== null && e.g !== '') nameDone[e.g] = (nameDone[e.g] || 0) + 1;
    });
    var pnOut = pnInput.map(function (p) {
      var done = nameDone[p.name] || 0;
      var t = (p.target === undefined || p.target === null) ? null : Number(p.target);
      var rate = (t && t > 0) ? done / t : 0;
      return { city: p.city, name: p.name, role: p.role, target: t, done: done, rate: rate };
    });

    /* ---------- 报告聚合 ---------- */
    var targetDone = pnOut.filter(function (p) { return p.target && p.target > 0 && p.done >= p.target; }).length;
    var zeroLeads = pnOut.filter(function (p) { return p.done === 0; }).length;
    var notDone = pnOut.filter(function (p) { return p.done > 0 && (!p.target || p.target <= 0 || p.done < p.target); }).length;
    var kpis = { totalLeads: entries.length, targetDone: targetDone, notDone: notDone, zeroLeads: zeroLeads, totalPeople: pnOut.length };

    var sorted = pnOut.slice().sort(function (a, b) {
      if (b.done !== a.done) return b.done - a.done;
      return String(a.name).localeCompare(String(b.name), 'zh');
    });
    var top8 = sorted.slice(0, 8).map(function (p) {
      return { name: p.name, leads: p.done };
    });

    var platCount = {};
    entries.forEach(function (e) {
      var kk = e.k === undefined || e.k === null ? '' : String(e.k);
      platCount[kk] = (platCount[kk] || 0) + 1;
    });
    var platList = Object.keys(platCount).sort(function (a, b) { return platCount[b] - platCount[a]; }).map(function (k) {
      return { platform: k, count: platCount[k] };
    });

    function aggBy(pickKey, outFn) {
      var map = {};
      pnOut.forEach(function (p) {
        var k = pickKey(p);
        if (!map[k]) map[k] = { key: k, people: 0, leads: 0, target: 0 };
        map[k].people++;
        map[k].leads += p.done;
        map[k].target += (p.target || 0);
      });
      var arr = Object.keys(map).map(function (k) {
        var a = map[k];
        a.rate = a.target > 0 ? a.leads / a.target : 0;
        return a;
      });
      arr.sort(function (a, b) { return b.leads - a.leads; });
      return arr;
    }
    var cityAgg = aggBy(function (p) {
      return (p.city === undefined || p.city === null) ? '(未填)' : String(p.city);
    });
    var roleAgg = aggBy(function (p) {
      return (p.role === undefined || p.role === null || String(p.role) === '') ? '(未填)' : String(p.role);
    });

    var cityOf = {}, roleOf = {};
    pnOut.forEach(function (p) { cityOf[p.name] = p.city; roleOf[p.name] = p.role; });
    function stat(src) { return src === 'water_only' ? SSY : (src === 'inter' ? SSY_SX : SX); }
    var intentRows = entries.map(function (e) {
      return {
        clue: e.f,
        city: e.g ? cityOf[e.g] : undefined,
        role: e.g ? roleOf[e.g] : undefined,
        owner: e.g,
        stat: stat(e.src),
        platform: e.k
      };
    });

    /* ---------- 生成输出工作簿 ---------- */
    var nwb = XLSX.utils.book_new();

    var crmSheet = XLSX.utils.aoa_to_sheet(cr);
    var sySheet = XLSX.utils.aoa_to_sheet(sy);
    var sxSheet = XLSX.utils.aoa_to_sheet(xr);

    var n1 = R1.length, n2 = entries.length, maxN = Math.max(n1, n2);
    var ppData = [];
    for (i = 0; i < maxN; i++) {
      var rA = ['', '', '', '', '', '', '', '', '', '', ''];
      if (i < n1) {
        var names = uniqueNames(ssyByPhone[R1[i]]);
        rA[0] = R1[i];
        for (var jj = 0; jj < Math.min(names.length, 3); jj++) rA[1 + jj] = names[jj];
        rA[4] = crmAct(R1[i]);
      }
      if (i < n2) {
        var e = entries[i];
        rA[5] = e.f; rA[6] = e.g; rA[9] = e.j; rA[10] = e.k;
      }
      ppData.push(rA);
    }
    var pp = XLSX.utils.aoa_to_sheet(ppData, { origin: 'A3' });
    XLSX.utils.sheet_add_aoa(pp, [['CRM与水师营-结果1', null, null, null, null, '结果1与线上表-结果2', null, null, null, null, null]], { origin: 'A1' });
    XLSX.utils.sheet_add_aoa(pp, [['线索匹配-1轮', '线索归属人', null, null, 'CRM活动', '线索匹配-2轮', '线索归属人-2轮', null, null, 'CRM活动', '平台来源']], { origin: 'A2' });
    pp['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
      { s: { r: 0, c: 5 }, e: { r: 0, c: 10 } },
      { s: { r: 1, c: 1 }, e: { r: 1, c: 3 } },
      { s: { r: 1, c: 6 }, e: { r: 1, c: 8 } }
    ];
    pp['!cols'] = [{ wch: 13 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 22 }, { wch: 13 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 22 }, { wch: 10 }];

    var ydAoa = [['线索', '分校', '岗位', '线索归属人', '统计来源', '平台来源']];
    intentRows.forEach(function (ir) {
      ydAoa.push([ir.clue, ir.city, ir.role, ir.owner, ir.stat, ir.platform]);
    });
    var yd = XLSX.utils.aoa_to_sheet(ydAoa);
    yd['!cols'] = [{ wch: 15 }, { wch: 14 }, { wch: 16 }, { wch: 12 }, { wch: 16 }, { wch: 10 }];

    var pnAoa = [['分校运营线索情况'], ['地市/学习中心\n全称', '姓名', '新分岗分类', '意向目标', '完成量', '完成率']];
    pnOut.forEach(function (p) {
      pnAoa.push([p.city, p.name, p.role, p.target, p.done, (p.rate * 100).toFixed(1) + '%']);
    });
    var pnSheet = XLSX.utils.aoa_to_sheet(pnAoa);
    pnSheet['!cols'] = [{ wch: 18 }, { wch: 10 }, { wch: 16 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];

    XLSX.utils.book_append_sheet(nwb, crmSheet, 'CRM');
    XLSX.utils.book_append_sheet(nwb, sySheet, '水师营');
    XLSX.utils.book_append_sheet(nwb, sxSheet, '线上表');
    XLSX.utils.book_append_sheet(nwb, pnSheet, '人员名单');
    XLSX.utils.book_append_sheet(nwb, yd, '意向明细');
    XLSX.utils.book_append_sheet(nwb, pp, '配对过程');

    /* ---------- 校验计数（供调试） ---------- */
    var waterOnly = 0, inter = 0, sxOnly = 0;
    entries.forEach(function (e) { if (e.src === 'water_only') waterOnly++; else if (e.src === 'inter') inter++; else sxOnly++; });
    var sxRealCount = Object.keys(sxRealMasked).length;
    var interReal = 0;
    Object.keys(sxRealMasked).forEach(function (f) { if (R1set.hasOwnProperty(f)) interReal++; });

    return {
      result1: R1,
      entries: entries,
      intentRows: intentRows,
      pnOut: pnOut,
      kpis: kpis,
      top8: top8,
      platList: platList,
      cityAgg: cityAgg,
      roleAgg: roleAgg,
      outputWb: nwb,
      counts: {
        result1: R1.length,
        entries: entries.length,
        waterOnly: waterOnly,
        inter: inter,
        sxOnly: sxOnly,
        interReal: interReal,
        sxRows: sxList.length,
        sxReal: sxRealCount,
        pnDoneTotal: pnOut.reduce(function (s, p) { return s + p.done; }, 0)
      }
    };
  }

  return { processWorkbook: processWorkbook };
});
