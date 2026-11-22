/*!
 * 青简叙 Export — 导出当前筛选结果
 *
 * CSV    : UTF-8 BOM，Excel / WPS 直接双击不乱码
 * Excel  : SheetJS，含「考试安排」「说明」两张工作表
 * PDF    : 逐页渲染 A4 画布再拼装，中文无需嵌入字体，手机端可用
 * Word   : Word/WPS 兼容 HTML 文档（.doc），保留表格与分页表头
 */
(function (root) {
  "use strict";

  var COLUMNS = [
    { key: "index", label: "序号", width: 6 },
    { key: "course", label: "课程名称", width: 26 },
    { key: "code", label: "课程号", width: 14 },
    { key: "date", label: "考试日期", width: 12 },
    { key: "weekday", label: "星期", width: 6 },
    { key: "start", label: "开始时间", width: 9 },
    { key: "end", label: "结束时间", width: 9 },
    { key: "room", label: "考试地点", width: 20 },
    { key: "grade", label: "年级", width: 8 },
    { key: "major", label: "专业", width: 10 },
    { key: "cls", label: "班级", width: 18 },
    { key: "headcount", label: "人数", width: 7 },
    { key: "college", label: "开课学院", width: 34 },
    { key: "note", label: "备注", width: 28 }
  ];

  var WEEK_CN = ["日", "一", "二", "三", "四", "五", "六"];

  function weekdayOf(date) {
    var d = new Date(date + "T00:00:00");
    return isNaN(d) ? "" : "周" + WEEK_CN[d.getDay()];
  }

  function pad(n) { return String(n).padStart(2, "0"); }

  function stamp() {
    var d = new Date();
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + "-" + pad(d.getHours()) + pad(d.getMinutes());
  }

  function timeText(rec) {
    return rec.start ? rec.start + (rec.end ? "-" + rec.end : "") : "待定";
  }

  function buildRows(records) {
    return records.map(function (r, i) {
      return [
        i + 1,
        r.course,
        r.code || "",
        r.date,
        weekdayOf(r.date),
        r.start || "",
        r.end || "",
        (r.rooms || []).join("、"),
        r.grade || "",
        r.major || "",
        r.cls || "多班合考",
        r.headcount == null ? "" : r.headcount,
        r.college || "",
        r.note || ""
      ];
    });
  }

  function download(blob, filename) {
    if (root.navigator && root.navigator.msSaveOrOpenBlob) {
      root.navigator.msSaveOrOpenBlob(blob, filename);
      return;
    }
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 1500);
  }

  function filenameBase(ctx) {
    var d = new Date();
    return (ctx.filenameBase || "考试安排") + "_" + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
  }

  function metaLines(ctx) {
    return [
      "总表：" + (ctx.datasetTitle || "—"),
      "范围：" + (ctx.scopeLabel || "全部"),
      "共 " + (ctx.count || 0) + " 场考试",
      "导出：" + new Date().toLocaleString("zh-CN", { hour12: false })
    ];
  }

  /* --------------------------------- CSV --------------------------------- */
  function toCSV(records, ctx) {
    var header = COLUMNS.map(function (c) { return c.label; });
    var lines = [header.join(",")];
    buildRows(records).forEach(function (row) {
      lines.push(row.map(function (cell) {
        var s = cell == null ? "" : String(cell);
        return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(","));
    });
    var blob = new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
    download(blob, filenameBase(ctx) + ".csv");
  }

  /* -------------------------------- Excel -------------------------------- */
  function toExcel(records, ctx) {
    var XLSX = root.XLSX;
    if (!XLSX) throw new Error("Excel 组件未加载，请刷新页面重试");

    var aoa = [COLUMNS.map(function (c) { return c.label; })].concat(buildRows(records));
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = COLUMNS.map(function (c) { return { wch: c.width }; });
    ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: COLUMNS.length - 1 } }) };
    ws["!freeze"] = { xSplit: "0", ySplit: "1" };

    var info = [["考试安排导出"], ["字段", "值"]];
    info.push(["总表", ctx.datasetTitle || ""]);
    info.push(["筛选范围", ctx.scopeLabel || "全部"]);
    info.push(["考试场次", records.length]);
    info.push(["课程门数", new Set(records.map(function (r) { return r.course; })).size]);
    info.push(["导出时间", new Date().toLocaleString("zh-CN", { hour12: false })]);
    info.push(["数据来源", "青简叙 · 考试纪事"]);
    var ws2 = XLSX.utils.aoa_to_sheet(info);
    ws2["!cols"] = [{ wch: 14 }, { wch: 46 }];

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "考试安排");
    XLSX.utils.book_append_sheet(wb, ws2, "说明");
    var out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    download(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filenameBase(ctx) + ".xlsx");
  }

  /* --------------------------------- Word -------------------------------- */
  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function toWord(records, ctx) {
    var rows = records.map(function (r, i) {
      return "<tr>" +
        "<td class='c'>" + (i + 1) + "</td>" +
        "<td>" + esc(r.course) + (r.note ? "<br/><span style='font-size:7.5pt;color:#777'>备注：" + esc(r.note) + "</span>" : "") + "</td>" +
        "<td class='c'>" + esc(r.date) + "<br/>" + weekdayOf(r.date) + "</td>" +
        "<td class='c'>" + esc(timeText(r)) + "</td>" +
        "<td>" + esc((r.rooms || []).join("、")) + "</td>" +
        "<td>" + esc(r.cls || "多班合考") + "</td>" +
        "</tr>";
    }).join("");

    var html = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>" +
      "<head><meta charset='utf-8'><title>" + esc(ctx.filenameBase || "考试安排") + "</title>" +
      "<style>" +
      "@page{size:A4 portrait;margin:1.6cm 1.2cm;}" +
      "body{font-family:'Microsoft YaHei','PingFang SC',sans-serif;font-size:10.5pt;color:#111;}" +
      "h1{font-size:17pt;margin:0 0 6pt;text-align:center;}" +
      ".sub{text-align:center;color:#555;font-size:9pt;margin-bottom:14pt;}" +
      "table{width:100%;border-collapse:collapse;font-size:9.5pt;}" +
      "th{background:#eef2f7;border:0.5pt solid #999;padding:5pt 4pt;font-weight:bold;}" +
      "td{border:0.5pt solid #bbb;padding:4pt;}" +
      "td.c{text-align:center;}" +
      "thead{display:table-header-group;}" +
      "tr{page-break-inside:avoid;}" +
      "</style></head><body>" +
      "<h1>" + esc(ctx.docTitle || "考试安排表") + "</h1>" +
      "<div class='sub'>" + metaLines(ctx).map(esc).join(" &nbsp;|&nbsp; ") + "</div>" +
      "<table><thead><tr><th style='width:6%'>#</th><th style='width:26%'>课程名称</th>" +
      "<th style='width:16%'>考试日期</th><th style='width:16%'>考试时间</th>" +
      "<th style='width:18%'>考试地点</th><th style='width:18%'>班级</th></tr></thead>" +
      "<tbody>" + rows + "</tbody></table>" +
      "<p style='font-size:8pt;color:#888;margin-top:12pt'>本表由「青简叙」导出，请以教务系统原始安排为准。</p>" +
      "</body></html>";

    download(new Blob(["\ufeff", html], { type: "application/msword;charset=utf-8" }), filenameBase(ctx) + ".doc");
  }

  /* --------------------------------- PDF --------------------------------- */
  var PAGE_W = 794;
  var PAGE_H = 1123;
  var ROWS_PER_PAGE = 24;

  function pageNode(records, ctx, pageNo, pageTotal) {
    var wrap = document.createElement("div");
    wrap.setAttribute("aria-hidden", "true");
    wrap.style.cssText = "position:fixed;left:-10000px;top:0;width:" + PAGE_W + "px;height:" + PAGE_H +
      "px;box-sizing:border-box;padding:34px 38px 46px;background:#fff;position:absolute;" +
      "font-family:'PingFang SC','Microsoft YaHei','Helvetica Neue',sans-serif;color:#0f172a;";

    var head = document.createElement("div");
    head.style.cssText = "display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #1e293b;padding-bottom:9px;";
    head.innerHTML =
      "<div><div style='font-size:19px;font-weight:700;letter-spacing:.5px'>" + esc(ctx.docTitle || "考试安排表") + "</div>" +
      "<div style='font-size:10px;color:#64748b;margin-top:5px'>" + esc(ctx.datasetTitle || "") + "</div></div>" +
      "<div style='text-align:right;font-size:9px;color:#64748b;line-height:1.75'>" +
      "<div>范围：" + esc(ctx.scopeLabel || "全部") + "</div>" +
      "<div>共 " + ctx.count + " 场 · 本页第 " + ((pageNo - 1) * ROWS_PER_PAGE + 1) + "–" +
      Math.min(pageNo * ROWS_PER_PAGE, ctx.count) + " 场</div>" +
      "<div>" + new Date().toLocaleDateString("zh-CN") + " 导出</div></div>";
    wrap.appendChild(head);

    var table = document.createElement("table");
    table.style.cssText = "width:100%;border-collapse:collapse;margin-top:10px;font-size:10.5px;table-layout:fixed;";
    table.innerHTML =
      "<thead><tr style='background:#eef2f7;color:#334155'>" +
      "<th style='width:30px;padding:6px 4px;border-bottom:1px solid #cbd5e1;text-align:center;font-weight:600'>#</th>" +
      "<th style='width:196px;padding:6px 6px;border-bottom:1px solid #cbd5e1;text-align:left;font-weight:600'>课程名称</th>" +
      "<th style='width:146px;padding:6px 6px;border-bottom:1px solid #cbd5e1;text-align:left;font-weight:600'>考试时间</th>" +
      "<th style='width:120px;padding:6px 6px;border-bottom:1px solid #cbd5e1;text-align:left;font-weight:600'>考试地点</th>" +
      "<th style='width:120px;padding:6px 6px;border-bottom:1px solid #cbd5e1;text-align:left;font-weight:600'>班级</th>" +
      "</tr></thead><tbody></tbody>";

    var tbody = table.querySelector("tbody");
    records.forEach(function (r) {
      var tr = document.createElement("tr");
      tr.style.cssText = "height:34px";
      var room = (r.rooms || []).join("、");
      tr.innerHTML =
        "<td style='text-align:center;color:#94a3b8;border-bottom:1px solid #eef2f7'>" + (r.__i) + "</td>" +
        "<td style='border-bottom:1px solid #eef2f7;padding:0 6px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>" + esc(r.course) + "</td>" +
        "<td style='border-bottom:1px solid #eef2f7;padding:0 6px;color:#334155'>" + esc(r.date) + " " + weekdayOf(r.date) + "<br/>" +
        "<span style='font-size:9px;color:#64748b'>" + esc(timeText(r)) + "</span></td>" +
        "<td style='border-bottom:1px solid #eef2f7;padding:0 6px;color:#334155;overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>" + esc(room) + "</td>" +
        "<td style='border-bottom:1px solid #eef2f7;padding:0 6px;color:#475569;font-size:9.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>" + esc(r.cls || "多班合考") + "</td>";
      tbody.appendChild(tr);
    });
    wrap.appendChild(table);

    var foot = document.createElement("div");
    foot.style.cssText = "position:absolute;left:38px;right:38px;bottom:22px;display:flex;justify-content:space-between;" +
      "font-size:8.5px;color:#94a3b8;border-top:1px solid #e2e8f0;padding-top:7px;";
    foot.innerHTML = "<span>青简叙 · 考试纪事　数据以教务系统为准</span><span>" + pageNo + " / " + pageTotal + "</span>";
    wrap.appendChild(foot);
    return wrap;
  }

  async function toPDF(records, ctx, onProgress) {
    var jsPDFCtor = root.jspdf && root.jspdf.jsPDF;
    if (!jsPDFCtor) throw new Error("PDF 组件未加载，请刷新页面重试");
    if (!root.html2canvas) throw new Error("PDF 渲染组件未加载，请刷新页面重试");

    var list = records.map(function (r, i) { return Object.assign({ __i: i + 1 }, r); });
    var total = Math.ceil(list.length / ROWS_PER_PAGE);
    var doc = new jsPDFCtor({ unit: "pt", format: "a4", compress: true, orientation: "portrait" });
    var pw = doc.internal.pageSize.getWidth();
    var ph = doc.internal.pageSize.getHeight();

    for (var p = 0; p < total; p++) {
      var slice = list.slice(p * ROWS_PER_PAGE, (p + 1) * ROWS_PER_PAGE);
      if (onProgress) onProgress(p + 1, total);
      var node = pageNode(slice, ctx, p + 1, total);
      document.body.appendChild(node);
      var canvas;
      try {
        canvas = await root.html2canvas(node, {
          backgroundColor: "#ffffff",
          scale: Math.min(2, (root.devicePixelRatio || 1) + 0.5),
          logging: false,
          useCORS: true,
          imageTimeout: 0
        });
      } finally {
        node.parentNode && node.parentNode.removeChild(node);
      }
      if (p > 0) doc.addPage();
      doc.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", 0, 0, pw, ph);
      await new Promise(function (r) { setTimeout(r, 0); });
    }
    doc.save(filenameBase(ctx) + ".pdf");
  }

  root.QingJianExport = {
    COLUMNS: COLUMNS,
    toCSV: toCSV,
    toExcel: toExcel,
    toWord: toWord,
    toPDF: toPDF,
    weekdayOf: weekdayOf,
    ROWS_PER_PAGE: ROWS_PER_PAGE
  };
})(typeof window !== "undefined" ? window : globalThis);
