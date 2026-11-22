/*!
 * 青简叙 Parser — 考试安排总表解析器
 * 同时支持浏览器（window.QingJianParser）与 Node（module.exports），
 * 供前端导入与 scripts/build-seed.mjs 复用同一套逻辑。
 *
 * 设计要点：
 * 1. 先用 !merges 把纵向合并单元格的真实值回填，而不是逐行 forward-fill。
 *    教务总表常出现「一个考试时间合并覆盖多个课程块」，forward-fill 会串行错配。
 * 2. 表头行通过别名打分自动定位，不假设固定在第 2 行。
 * 3. 年级 / 专业 / 班号从「24级汉语言01班」中拆解，兼容独立列写法。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.QingJianParser = factory();
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  /* ------------------------------------------------------------------ *
   * 字段别名表：score 越高越优先，避免「教学班组成」被误判为「班级」
   * ------------------------------------------------------------------ */
  var FIELDS = [
    { key: "course", aliases: ["课程名称", "考试课程", "课程名", "考试科目", "科目", "课程"] },
    { key: "code", aliases: ["课程号", "课程代码", "课程编号"] },
    { key: "date", aliases: ["考试日期", "日期"] },
    { key: "time", aliases: ["考试时间", "考试时段", "时间段", "时间", "时段"] },
    { key: "room", aliases: ["已排地点", "考试地点", "考场", "考试教室", "地点", "教室"] },
    { key: "cls", aliases: ["班级", "行政班", "班级名称", "学生班级", "班级简称"] },
    { key: "grade", aliases: ["年级", "级"] },
    { key: "major", aliases: ["专业", "专业名称"] },
    { key: "college", aliases: ["开课学院", "开课单位", "学院", "开课部门"] },
    { key: "count", aliases: ["人数", "报考人数", "考试人数", "考生人数"] },
    { key: "week", aliases: ["周次", "教学周"] },
    { key: "weekday", aliases: ["星期", "星期几", "周几"] },
    { key: "teaching", aliases: ["教学班组成", "教学班", "合班", "考试班级"] }
  ];

  var SKIP_ROW = /^(合计|小计|总计|备注|说明|注[:：]|制表|审核|注$)/;

  /* ------------------------------ 文本工具 ------------------------------ */
  function normText(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\u3000/g, " ").replace(/\s+/g, " ").trim();
  }

  function normHeader(v) {
    return normText(v)
      .replace(/[（(][^）)]*[）)]/g, "")
      .replace(/[\s:：*＊]/g, "")
      .toLowerCase();
  }

  function firstText() {
    for (var i = 0; i < arguments.length; i++) {
      var t = normText(arguments[i]);
      if (t) return t;
    }
    return "";
  }

  /* --------------------------- 表头行自动定位 --------------------------- */
  function scoreRow(row) {
    var map = {};
    var matched = 0;
    var weight = 0;

    for (var c = 0; c < row.length; c++) {
      var h = normHeader(row[c]);
      if (!h || h.length > 16) continue;

      var best = null;
      for (var f = 0; f < FIELDS.length; f++) {
        var field = FIELDS[f];
        if (map[field.key] !== undefined) continue;
        for (var a = 0; a < field.aliases.length; a++) {
          var alias = field.aliases[a];
          var score = 0;
          if (h === alias) score = 100;
          else if (h.indexOf(alias) === 0) score = 70;
          else if (h.indexOf(alias) > 0) score = 40;
          else continue;
          // 别名越长越可信：避免「课程」抢走「课程号」
          score += alias.length;
          if (!best || score > best.score) best = { key: field.key, score: score };
        }
      }
      if (best) {
        map[best.key] = c;
        matched++;
        weight += best.score;
      }
    }
    return { map: map, matched: matched, weight: weight };
  }

  function detectHeader(matrix) {
    var limit = Math.min(matrix.length, 16);
    var best = null;
    for (var r = 0; r < limit; r++) {
      var s = scoreRow(matrix[r] || []);
      var ok = s.map.course !== undefined && (s.map.time !== undefined || s.map.date !== undefined);
      if (!ok) continue;
      var hasOwner = s.map.cls !== undefined || s.map.grade !== undefined || s.map.major !== undefined;
      if (!hasOwner) continue;
      if (!best || s.weight > best.weight) best = { row: r, map: s.map, weight: s.weight };
    }
    return best;
  }

  /* ------------------------------ 矩阵构建 ------------------------------ */
  function buildMatrix(ws, XLSX) {
    if (!ws || !ws["!ref"]) return null;
    var range = XLSX.utils.decode_range(ws["!ref"]);
    var rows = range.e.r + 1;
    var cols = range.e.c + 1;
    var matrix = [];

    for (var r = 0; r < rows; r++) {
      var line = new Array(cols);
      for (var c = 0; c < cols; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        line[c] = cell && cell.v !== undefined ? cell.v : null;
      }
      matrix.push(line);
    }

    // 合并区域回填：把左上角的值广播到整个合并区
    var merges = ws["!merges"] || [];
    for (var m = 0; m < merges.length; m++) {
      var mg = merges[m];
      var v = matrix[mg.s.r] ? matrix[mg.s.r][mg.s.c] : null;
      if (v === null || v === undefined || normText(v) === "") continue;
      for (var mr = mg.s.r; mr <= mg.e.r; mr++) {
        if (!matrix[mr]) continue;
        for (var mc = mg.s.c; mc <= mg.e.c; mc++) {
          if (mc < matrix[mr].length) matrix[mr][mc] = v;
        }
      }
    }
    return matrix;
  }

  /* ------------------------------ 时间解析 ------------------------------ */
  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function serialToParts(n) {
    // Excel 1900 日期系统：25569 = 1970-01-01
    var ms = Math.round((n - 25569) * 86400000);
    var d = new Date(ms);
    return { date: d, hh: d.getUTCHours(), mm: d.getUTCMinutes() };
  }

  function normalizeHour(h, period) {
    var hour = h;
    if (/下午|晚上|傍晚/.test(period || "") && hour < 12) hour += 12;
    if (/中午/.test(period || "") && hour < 11) hour += 12;
    return hour;
  }

  /**
   * 支持：2026-11-20(19:30-21:10) / 2026/11/20 19:30-21:10 /
   *       11月20日 晚上7:30 / Excel 序列号 / Date 对象
   */
  function parseExamTime(timeRaw, dateRaw, yearHint) {
    var out = { date: "", start: "", end: "" };
    var timeText = normText(timeRaw);
    var dateText = normText(dateRaw);
    var year = yearHint || new Date().getFullYear();

    // 1. 纯日期列 + 纯时间列
    if (timeText && !/\d{4}[-/年]/.test(timeText) && dateText) {
      var u = timeText;
      if (/^\d{5}(\.\d+)?$/.test(u)) {
        var parts = serialToParts(parseFloat(u));
        return {
          date: parts.date.getUTCFullYear() + "-" + pad2(parts.date.getUTCMonth() + 1) + "-" + pad2(parts.date.getUTCDate()),
          start: pad2(parts.hh) + ":" + pad2(parts.mm),
          end: ""
        };
      }
      var rangeOnly = u.match(/(上午|下午|晚上|中午)?\s*(\d{1,2})\s*[:：]\s*(\d{2})\s*[-–—~～至]\s*(上午|下午|晚上|中午)?\s*(\d{1,2})\s*[:：]\s*(\d{2})/);
      if (rangeOnly) {
        out.start = pad2(normalizeHour(+rangeOnly[2], rangeOnly[1])) + ":" + rangeOnly[3];
        out.end = pad2(normalizeHour(+rangeOnly[5], rangeOnly[4])) + ":" + rangeOnly[6];
      }
      timeText = dateText + " " + u;
    }

    if (!timeText && !dateText) return out;

    var seq = /^\d{5}(\.\d+)?$/.test(timeText) ? parseFloat(timeText) : null;
    if (seq !== null && seq > 20000) {
      var p = serialToParts(seq);
      out.date = p.date.getUTCFullYear() + "-" + pad2(p.date.getUTCMonth() + 1) + "-" + pad2(p.date.getUTCDate());
      out.start = pad2(p.hh) + ":" + pad2(p.mm);
      return out;
    }

    if (timeRaw instanceof Date && !isNaN(timeRaw)) {
      var dt = timeRaw;
      out.date = dt.getFullYear() + "-" + pad2(dt.getMonth() + 1) + "-" + pad2(dt.getDate());
      if (dt.getHours() || dt.getMinutes()) out.start = pad2(dt.getHours()) + ":" + pad2(dt.getMinutes());
      return out;
    }

    // 2. 日期
    var dm = timeText.match(/(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?/);
    if (dm) {
      out.date = dm[1] + "-" + pad2(dm[2]) + "-" + pad2(dm[3]);
    } else {
      var m2 = timeText.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*日?/);
      if (m2) out.date = year + "-" + pad2(m2[1]) + "-" + pad2(m2[2]);
    }

    // 3. 起止时间
    var tm = timeText.match(/(上午|下午|晚上|中午)?\s*(\d{1,2})\s*[:：]\s*(\d{2})\s*[-–—~～至到]\s*(上午|下午|晚上|中午)?\s*(\d{1,2})\s*[:：]\s*(\d{2})/);
    if (tm) {
      out.start = pad2(normalizeHour(+tm[2], tm[1])) + ":" + tm[3];
      out.end = pad2(normalizeHour(+tm[5], tm[4])) + ":" + tm[6];
    } else {
      var single = timeText.match(/(上午|下午|晚上|中午)?\s*(\d{1,2})\s*[:：]\s*(\d{2})/);
      if (single) out.start = pad2(normalizeHour(+single[2], single[1])) + ":" + single[3];
    }

    return out;
  }

  /* ---------------------------- 班级字段拆解 ---------------------------- */
  function parseClassLabel(raw, gradeCol, majorCol) {
    var s = normText(raw).replace(/\s+/g, "");
    var result = { grade: normText(gradeCol), major: normText(majorCol), classNo: "", cls: s };

    if (s) {
      var m = s.match(/^(\d{2,4})\s*级\s*(.+?)\s*(\d{1,3})\s*班$/);
      if (m) {
        var yy = m[1].length === 4 ? m[1].slice(2) : m[1];
        if (!result.grade) result.grade = yy + "级";
        if (!result.major) result.major = m[2];
        result.classNo = m[3].padStart(2, "0");
        result.cls = result.grade + result.major + result.classNo + "班";
        return result;
      }
      var m2 = s.match(/^(.+?)\s*(\d{1,3})\s*班$/);
      if (m2) {
        if (!result.major) result.major = m2[1];
        result.classNo = m2[2].padStart(2, "0");
      }
    }

    if (result.grade) {
      var g = result.grade.match(/(\d{2,4})\s*级?/);
      if (g) {
        var y2 = g[1].length === 4 ? g[1].slice(2) : g[1];
        result.grade = y2 + "级";
      }
    }
    return result;
  }

  function splitList(raw) {
    return normText(raw)
      .split(/[,，;；、/|]/)
      .map(function (x) { return x.trim(); })
      .filter(Boolean)
      .filter(function (x, i, arr) { return arr.indexOf(x) === i; });
  }

  /* ------------------------------- 主流程 ------------------------------- */
  function parseSheet(ws, sheetName, XLSX, opts) {
    opts = opts || {};
    var matrix = buildMatrix(ws, XLSX);
    if (!matrix) return null;

    var header = detectHeader(matrix);
    if (!header) return null;

    var map = header.map;
    var records = [];
    var skipped = 0;
    var inherited = 0;
    var lastCourse = null;
    var lastCode = null;
    var lastCollege = null;
    var yearHint = opts.yearHint || new Date().getFullYear();

    function cell(row, key) {
      var idx = map[key];
      if (idx === undefined) return null;
      return row[idx];
    }

    for (var r = header.row + 1; r < matrix.length; r++) {
      var row = matrix[r] || [];

      // 整行空白
      var hasAny = row.some(function (v) { return normText(v) !== ""; });
      if (!hasAny) continue;

      var rawCourse = normText(cell(row, "course"));
      if (SKIP_ROW.test(rawCourse)) { skipped++; continue; }

      var clsInfo = parseClassLabel(cell(row, "cls"), cell(row, "grade"), cell(row, "major"));
      var timeInfo = parseExamTime(cell(row, "time"), cell(row, "date"), yearHint);

      // 无课程且无班级 → 非数据行
      if (!rawCourse && !clsInfo.cls) { skipped++; continue; }
      // 只有课程、没有班级也没有时间 → 标题类行
      if (!clsInfo.cls && !timeInfo.date) { skipped++; continue; }

      var course = rawCourse || lastCourse;
      if (!rawCourse && course) inherited++;
      if (rawCourse) lastCourse = rawCourse;

      var code = firstText(cell(row, "code"), lastCode);
      if (normText(cell(row, "code"))) lastCode = normText(cell(row, "code"));

      var college = firstText(cell(row, "college"), lastCollege);
      if (normText(cell(row, "college"))) lastCollege = normText(cell(row, "college"));

      if (!course) { skipped++; continue; }
      if (!timeInfo.date) { skipped++; continue; }

      var rooms = splitList(cell(row, "room"));
      var teaching = splitList(cell(row, "teaching"));
      var count = parseInt(normText(cell(row, "count")).replace(/[^\d]/g, ""), 10);

      var classKey = clsInfo.cls ? [clsInfo.grade, clsInfo.major, clsInfo.classNo].join("|") : "";

      records.push({
        id: "",
        course: course,
        code: code,
        college: college,
        date: timeInfo.date,
        start: timeInfo.start,
        end: timeInfo.end,
        week: parseInt(normText(cell(row, "week")), 10) || null,
        weekday: normText(cell(row, "weekday")),
        grade: clsInfo.grade,
        major: clsInfo.major,
        classNo: clsInfo.classNo,
        cls: clsInfo.cls,
        classKey: classKey,
        headcount: isNaN(count) ? null : count,
        rooms: rooms,
        teaching: teaching.join(";"),
        sheet: sheetName,
        row: r + 1
      });
    }

    if (!records.length) return null;

    return {
      sheet: sheetName,
      headerRow: header.row + 1,
      columns: map,
      records: records,
      stats: { inherited: inherited, skipped: skipped }
    };
  }

  function dayStamp(dateStr, timeStr) {
    if (!dateStr) return 0;
    var t = timeStr ? timeStr : "23:59";
    var ms = Date.parse(dateStr + "T" + t + ":00");
    return isNaN(ms) ? Date.parse(dateStr + "T23:59:00") : ms;
  }

  function addMinutes(hhmm, mins) {
    var p = String(hhmm || "").split(":");
    if (p.length < 2) return "";
    var total = (+p[0]) * 60 + (+p[1]) + mins;
    total = ((total % 1440) + 1440) % 1440;
    return pad2(Math.floor(total / 60)) + ":" + pad2(total % 60);
  }

  function finalizeRecords(records, datasetId) {
    records.sort(function (a, b) {
      var d = (a.date + a.start).localeCompare(b.date + b.start);
      if (d) return d;
      if (a.cls !== b.cls) return a.cls.localeCompare(b.cls, "zh");
      return a.course.localeCompare(b.course, "zh");
    });
    records.forEach(function (r, i) {
      r.id = datasetId + "#" + (i + 1);
      // 缺失结束时间时按 110 分钟（标准考试时长）推算，避免倒计时把「正在考试」误判为已结束
      var endTime = r.end || (r.start ? addMinutes(r.start, 110) : "23:59");
      r.end = r.end || endTime;
      r.iso = r.date + "T" + endTime + ":00";
    });
    return records;
  }

  function summarize(records) {
    var courses = new Set();
    var classes = new Set();
    var grades = new Set();
    var majors = new Set();
    var dates = [];
    records.forEach(function (r) {
      courses.add(r.course);
      if (r.cls) classes.add(r.cls);
      if (r.grade) grades.add(r.grade);
      if (r.major) majors.add(r.major);
      dates.push(r.date);
    });
    dates.sort();
    return {
      recordCount: records.length,
      courseCount: courses.size,
      classCount: classes.size,
      grades: Array.from(grades).sort(),
      majors: Array.from(majors).sort(function (a, b) { return a.localeCompare(b, "zh"); }),
      classes: Array.from(classes).sort(function (a, b) { return a.localeCompare(b, "zh"); }),
      dateFrom: dates[0] || "",
      dateTo: dates[dates.length - 1] || ""
    };
  }

  /**
   * 解析一个工作簿。多个 sheet 时选记录数最多且表头最完整的那张。
   * @returns {{ok:boolean, dataset?:object, error?:string, warnings?:string[]}}
   */
  function parseWorkbook(wb, XLSX, opts) {
    opts = opts || {};
    var candidates = [];
    var warnings = [];

    (wb.SheetNames || []).forEach(function (name) {
      var ws = wb.Sheets[name];
      if (!ws) return;
      var parsed = parseSheet(ws, name, XLSX, opts);
      if (parsed) candidates.push(parsed);
    });

    if (!candidates.length) {
      return {
        ok: false,
        error: "未识别到考试安排表头。请确认表格包含「课程名称」「考试时间」「班级」等列。",
        warnings: warnings
      };
    }

    candidates.sort(function (a, b) { return b.records.length - a.records.length; });
    var chosen = candidates[0];

    if (candidates.length > 1) {
      warnings.push(
        "该文件含 " + candidates.length + " 张可解析工作表，已选用数据量最大的「" +
        chosen.sheet + "」（" + chosen.records.length + " 条）。"
      );
    }
    if (chosen.stats.inherited > 0) {
      warnings.push(chosen.stats.inherited + " 条记录通过向上继承补齐了课程/时间列（原表为视觉合并但未设合并区域）。");
    }
    if (chosen.stats.skipped > 0) {
      warnings.push("已跳过 " + chosen.stats.skipped + " 行非数据内容（标题、合计或空行）。");
    }

    var records = chosen.records;
    var noClass = records.filter(function (r) { return !r.cls; }).length;
    if (noClass) warnings.push(noClass + " 条记录没有具体班级（如多班合考），仅能按年级/专业检索。");

    var noRoom = records.filter(function (r) { return !r.rooms.length; }).length;
    if (noRoom) warnings.push(noRoom + " 条记录缺少考试地点。");

    // 教务表常见「补录行」：班级列写 A 班，教学班组成却是 B 班，人数为 1。
    // 以班级列为筛选依据，但必须提示，避免同学误以为数据错乱。
    var mismatch = records.filter(function (r) {
      if (!r.cls || !r.teaching) return false;
      return r.teaching.split(";").indexOf(r.cls) === -1;
    }).length;
    if (mismatch) {
      warnings.push(mismatch + " 条记录的「班级」不在同行「教学班组成」中（多为教务补录或合考行），筛选以班级列为准。");
    }

    var id = opts.datasetId || "ds-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    finalizeRecords(records, id);

    var summary = summarize(records);
    var title = opts.title || "";
    if (!title) {
      // 优先用表格标题横幅，其次用课程 + 日期范围描述
      var banner = "";
      for (var r = 0; r < Math.min(chosen.headerRow - 1, 3); r++) {
        var line = (buildSheetFirstCells(wb, XLSX, chosen.sheet, r) || []).join(" ");
        if (/考试|安排/.test(line) && line.length > banner.length) banner = line.trim();
      }
      title = banner || (summary.dateFrom ? summary.dateFrom.slice(0, 4) + " 学年考试安排" : "考试安排");
    }

    return {
      ok: true,
      warnings: warnings,
      dataset: {
        id: id,
        title: title,
        sourceName: opts.sourceName || "",
        importedAt: new Date().toISOString(),
        sheet: chosen.sheet,
        headerRow: chosen.headerRow,
        columns: chosen.columns,
        summary: summary,
        warnings: warnings
      },
      records: records
    };
  }

  function buildSheetFirstCells(wb, XLSX, sheetName, rowIdx) {
    var ws = wb.Sheets[sheetName];
    if (!ws || !ws["!ref"]) return [];
    var range = XLSX.utils.decode_range(ws["!ref"]);
    var cells = [];
    for (var c = range.s.c; c <= Math.min(range.e.c, 12); c++) {
      var cell = ws[XLSX.utils.encode_cell({ r: rowIdx, c: c })];
      if (cell && cell.v) cells.push(String(cell.v));
    }
    return cells;
  }

  return {
    parseWorkbook: parseWorkbook,
    parseSheet: parseSheet,
    parseExamTime: parseExamTime,
    parseClassLabel: parseClassLabel,
    buildMatrix: buildMatrix,
    detectHeader: detectHeader,
    summarize: summarize,
    dayStamp: dayStamp,
    splitList: splitList
  };
});
