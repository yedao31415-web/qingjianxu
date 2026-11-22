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
    { key: "course", aliases: ["课程名称", "考试课程", "课程名", "考试科目", "科目", "课程", "coursename", "course", "subject"] },
    { key: "code", aliases: ["课程号", "课程代码", "课程编号", "coursecode", "code"] },
    { key: "date", aliases: ["考试日期", "日期", "examdate", "date"] },
    { key: "time", aliases: ["考试时间", "考试时段", "时间段", "时间", "时段", "examtime"] },
    { key: "startTime", aliases: ["开始时间", "开考时间", "起始时间", "考试开始时间", "starttime", "start"] },
    { key: "endTime", aliases: ["结束时间", "终止时间", "考试结束时间", "endtime", "end"] },
    { key: "room", aliases: ["已排地点", "考试地点", "考场", "考试教室", "地点", "教室", "课室", "试室", "venue", "room", "location"] },
    { key: "cls", aliases: ["班级", "行政班", "班级名称", "学生班级", "班级简称", "classname", "class"] },
    { key: "grade", aliases: ["年级", "级", "grade"] },
    { key: "major", aliases: ["专业", "专业名称", "major"] },
    { key: "college", aliases: ["开课学院", "开课单位", "学院", "开课部门", "院系", "系部", "college", "department"] },
    { key: "count", aliases: ["人数", "报考人数", "考试人数", "考生人数", "考生数", "count", "students"] },
    { key: "week", aliases: ["周次", "教学周", "week"] },
    { key: "weekday", aliases: ["星期", "星期几", "周几", "weekday"] },
    { key: "teaching", aliases: ["教学班组成", "教学班", "合班", "考试班级", "考试对象"] }
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
      var ok = s.map.course !== undefined && (s.map.time !== undefined || s.map.date !== undefined || s.map.startTime !== undefined);
      if (!ok) continue;
      // 只写「考试班级 / 教学班」的表也能用：后续会按班级拆行
      var hasOwner =
        s.map.cls !== undefined ||
        s.map.grade !== undefined ||
        s.map.major !== undefined ||
        s.map.teaching !== undefined;
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

  function fmtDate(d) {
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }

  /** 解析日期文本；缺年份时按“最近考期”推断，避免 1 月表被算成上一年 */
  function parseDateText(text, year) {
    var t = normText(text);
    if (!t) return "";
    var dm = t.match(/(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?/);
    if (dm) return dm[1] + "-" + pad2(dm[2]) + "-" + pad2(dm[3]);
    var m2 = t.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*日?/);
    if (m2) {
      var y = Number(year) || new Date().getFullYear();
      var guess = new Date(y, +m2[1] - 1, +m2[2]);
      if (guess.getTime() < Date.now() - 180 * 86400000) y += 1;
      return y + "-" + pad2(m2[1]) + "-" + pad2(m2[2]);
    }
    return "";
  }

  /** 解析纯时段：09:00-11:00 / 上午9:00—11:00 / 9点 / Excel 时间小数 */
  function clockRange(raw) {
    var u = normText(raw);
    var out = { start: "", end: "" };
    if (!u) return out;

    var rm = u.match(/(上午|下午|晚上|中午)?\s*(\d{1,2})\s*[:：]\s*(\d{2})\s*[-–—~～至到]\s*(上午|下午|晚上|中午)?\s*(\d{1,2})\s*[:：]\s*(\d{2})/);
    if (rm) {
      out.start = pad2(normalizeHour(+rm[2], rm[1])) + ":" + rm[3];
      out.end = pad2(normalizeHour(+rm[5], rm[4])) + ":" + rm[6];
      return out;
    }

    var sm = u.match(/(上午|下午|晚上|中午)?\s*(\d{1,2})\s*[:：]\s*(\d{2})/);
    if (sm) {
      out.start = pad2(normalizeHour(+sm[2], sm[1])) + ":" + sm[3];
      return out;
    }

    var cm = u.match(/(上午|下午|晚上|中午)?\s*(\d{1,2})\s*点\s*(半|\d{1,2})?/);
    if (cm) {
      var mm = cm[3] === "半" ? 30 : (cm[3] ? +cm[3] : 0);
      out.start = pad2(normalizeHour(+cm[2], cm[1])) + ":" + pad2(mm);
      return out;
    }

    if (/^0?\.\d+$/.test(u)) {
      var total = Math.round(parseFloat(u) * 1440);
      total = ((total % 1440) + 1440) % 1440;
      out.start = pad2(Math.floor(total / 60)) + ":" + pad2(total % 60);
    }
    return out;
  }

  /** 日期列统一成文本：支持文本、Date、Excel 序列号 */
  function normalizeDateRaw(raw) {
    if (raw instanceof Date && !isNaN(raw)) return fmtDate(raw);
    var t = normText(raw);
    if (/^\d{5}(\.\d+)?$/.test(t)) {
      var p = serialToParts(parseFloat(t));
      return p.date.getUTCFullYear() + "-" + pad2(p.date.getUTCMonth() + 1) + "-" + pad2(p.date.getUTCDate());
    }
    return t;
  }

  /**
   * 支持：2026-11-20(19:30-21:10) / 2026/11/20 19:30-21:10 /
   *       11月20日 晚上7:30 / Excel 序列号 / Date 对象 /
   *       日期与时间分列 / 日期列是 Excel 日期 / 只有“第 N 场”
   */
  function parseExamTime(timeRaw, dateRaw, yearHint) {
    var out = { date: "", start: "", end: "" };
    var year = yearHint || new Date().getFullYear();
    var timeText = normText(timeRaw);
    var dateText = normalizeDateRaw(dateRaw);

    if (timeRaw instanceof Date && !isNaN(timeRaw)) {
      out.date = fmtDate(timeRaw);
      if (timeRaw.getHours() || timeRaw.getMinutes()) {
        out.start = pad2(timeRaw.getHours()) + ":" + pad2(timeRaw.getMinutes());
      }
      if (!out.date && dateText) out.date = parseDateText(dateText, year);
      return out;
    }

    if (timeText && /^\d+(\.\d+)?$/.test(timeText)) {
      var num = parseFloat(timeText);
      if (num > 20000) {
        var p = serialToParts(num);
        out.date = p.date.getUTCFullYear() + "-" + pad2(p.date.getUTCMonth() + 1) + "-" + pad2(p.date.getUTCDate());
        out.start = pad2(p.hh) + ":" + pad2(p.mm);
        if (!out.date && dateText) out.date = parseDateText(dateText, year);
        return out;
      }
      if (num > 0 && num < 1) {
        out.start = clockRange(timeText).start;
        out.date = parseDateText(dateText, year);
        return out;
      }
    }

    out.date = parseDateText(timeText, year) || parseDateText(dateText, year);
    var clock = clockRange(timeText);
    if (!clock.start) clock = clockRange(dateText);
    out.start = clock.start;
    out.end = clock.end;
    return out;
  }

  /* ---------------------------- 班级字段拆解 ---------------------------- */
  function normalizeGrade(g) {
    var t = normText(g);
    var m = t.match(/(\d{2,4})\s*级?/);
    if (!m) return t;
    var y = m[1].length === 4 ? m[1].slice(2) : m[1];
    return y + "级";
  }

  /**
   * 兼容常见班级写法：
   *   2024级计算机科学与技术（1）班 / 24级计科1班
   *   2024软件工程1班 / 计科2401班 / 软件24-3班 / 24计科1班
   *   软件工程（1）班 / 2401班
   * 年级、专业若另有独立列，以独立列为准。
   */
  function parseClassLabel(raw, gradeCol, majorCol) {
    var original = normText(raw);
    var s = original.replace(/\s+/g, "").replace(/（/g, "(").replace(/）/g, ")");
    var result = { grade: normalizeGrade(gradeCol), major: normText(majorCol), classNo: "", cls: original };
    if (!s) return result;

    var m;
    var grade = "";
    var major = "";
    var classNo = "";

    if ((m = s.match(/^(\d{2,4})级(.+?)(?:\((\d{1,3})\)|(\d{1,3}))班$/))) {
      grade = m[1]; major = m[2]; classNo = m[3] || m[4];
    } else if ((m = s.match(/^(\d{4})([^\d].*?)(?:\((\d{1,3})\)|(\d{1,3}))班$/))) {
      grade = m[1]; major = m[2]; classNo = m[3] || m[4];
    } else if ((m = s.match(/^([^\d].*?)(\d{4})班$/))) {
      major = m[1]; grade = m[2].slice(0, 2); classNo = m[2].slice(2);
    } else if ((m = s.match(/^(.+?)(\d{2})[-—](\d{1,3})班$/))) {
      major = m[1]; grade = m[2]; classNo = m[3];
    } else if ((m = s.match(/^(\d{2})([^\d].*?)(?:\((\d{1,3})\)|(\d{1,3}))班$/))) {
      grade = m[1]; major = m[2]; classNo = m[3] || m[4];
    } else if ((m = s.match(/^(\d{1,4})班$/))) {
      if (m[1].length === 4) { grade = m[1].slice(0, 2); classNo = m[1].slice(2); }
      else classNo = m[1];
    } else if ((m = s.match(/^(.+?)(?:\((\d{1,3})\)|(\d{1,3}))班$/))) {
      major = m[1]; classNo = m[2] || m[3];
    } else if ((m = s.match(/^(\d{2,4})级(.+)$/))) {
      grade = m[1]; major = m[2];
    } else if ((m = s.match(/^([^\d].*?)(\d{4})$/))) {
      major = m[1]; grade = m[2].slice(0, 2); classNo = m[2].slice(2);
    }

    if (classNo) result.classNo = String(classNo).padStart(2, "0");
    if (!result.grade && grade) result.grade = normalizeGrade(grade);
    if (!result.major && major) result.major = major;

    if (result.grade || result.major || result.classNo) {
      result.cls = (result.grade || "") + (result.major || "") + (result.classNo || "") + "班";
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

  function weekdayOf(dateStr) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateStr || "");
    if (!m) return "";
    var d = new Date(+m[1], +m[2] - 1, +m[3]);
    return "日一二三四五六".charAt(d.getDay());
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

      var timeInfo = parseExamTime(cell(row, "time"), cell(row, "date"), yearHint);

      // 分列的开始时间 / 结束时间（不少学校把时段拆成两列）
      if (!timeInfo.start && (cell(row, "startTime") || cell(row, "endTime"))) {
        var startClock = clockRange(cell(row, "startTime"));
        var endClock = clockRange(cell(row, "endTime"));
        timeInfo.start = startClock.start || startClock.end || "";
        timeInfo.end = endClock.start || endClock.end || "";
      }

      // 班级：优先「班级」列；只有「考试班级 / 教学班」时按班级拆成多行
      var clsLabels = splitList(cell(row, "cls"));
      var teachLabels = splitList(cell(row, "teaching"));
      var labels = clsLabels.length ? clsLabels : teachLabels;
      var gradeCol = normText(cell(row, "grade"));
      var majorCol = normText(cell(row, "major"));
      var hasOwner = labels.length > 0 || !!gradeCol || !!majorCol;

      if (!rawCourse && !hasOwner) { skipped++; continue; }
      if (!hasOwner && !timeInfo.date) { skipped++; continue; }

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
      var count = parseInt(normText(cell(row, "count")).replace(/[^\d]/g, ""), 10);
      var weekdayText = normText(cell(row, "weekday")) || weekdayOf(timeInfo.date);
      var teachingText = teachLabels.length
        ? teachLabels.join(";")
        : (labels.length > 1 ? labels.join(";") : "");
      if (!labels.length) labels = [""];
      // 一行覆盖多个班时，人数是合计值，拆行后按班数均分
      var perCount = labels.length > 1 && !isNaN(count)
        ? Math.max(1, Math.round(count / labels.length))
        : count;

      labels.forEach(function (label) {
        var clsInfo = parseClassLabel(label, gradeCol, majorCol);
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
          weekday: weekdayText,
          grade: clsInfo.grade,
          major: clsInfo.major,
          classNo: clsInfo.classNo,
          cls: clsInfo.cls,
          classKey: classKey,
          headcount: isNaN(perCount) ? null : perCount,
          rooms: rooms,
          teaching: teachingText,
          sheet: sheetName,
          row: r + 1
        });
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
   * 解析一个工作簿。多个 sheet 会合并导入（按课程/日期/时段/班级/地点去重）。
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

    // 多工作表（常见于按学院/年级分表）合并导入，完全相同的行去重
    var records = [];
    var seen = {};
    var duplicated = 0;
    var inheritedTotal = 0;
    var skippedTotal = 0;
    candidates.forEach(function (c) {
      inheritedTotal += c.stats.inherited;
      skippedTotal += c.stats.skipped;
      c.records.forEach(function (rec) {
        var key = [rec.course, rec.date, rec.start, rec.end, rec.cls, (rec.rooms || []).join(","), rec.teaching].join("|");
        if (seen[key]) { duplicated++; return; }
        seen[key] = 1;
        records.push(rec);
      });
    });

    if (candidates.length > 1) {
      warnings.push(
        "该文件含 " + candidates.length + " 张可解析工作表（" +
        candidates.map(function (c) { return c.sheet; }).slice(0, 4).join("、") +
        (candidates.length > 4 ? " 等" : "") + "），已合并 " + records.length + " 条记录" +
        (duplicated ? "，忽略重复 " + duplicated + " 条" : "") + "。"
      );
    }
    if (inheritedTotal > 0) {
      warnings.push(inheritedTotal + " 条记录通过向上继承补齐了课程/时间列（原表为视觉合并但未设合并区域）。");
    }
    if (skippedTotal > 0) {
      warnings.push("已跳过 " + skippedTotal + " 行非数据内容（标题、合计或空行）。");
    }
    var noClass = records.filter(function (r) { return !r.cls; }).length;
    if (noClass) warnings.push(noClass + " 条记录没有具体班级（如多班合考），仅能按年级/专业检索。");

    var noRoom = records.filter(function (r) { return !r.rooms.length; }).length;
    if (noRoom) warnings.push(noRoom + " 条记录缺少考试地点。");

    var noTime = records.filter(function (r) { return !r.start; }).length;
    if (noTime) warnings.push(noTime + " 条记录只有日期、没有具体时段，倒数按当天结束计算。");

    // 教务表常见「补录行」：班级列写 A 班，教学班组成却是 B 班，人数为 1。
    // 以班级列为筛选依据，但必须提示，避免同学误以为数据错乱。
    var mismatch = records.filter(function (r) {
      if (!r.cls || !r.teaching || !r.classKey) return false;
      // 班级写法可能一处带年级、一处不带，按解析后的年级/专业/班号比较
      var keys = r.teaching.split(";").map(function (label) {
        var info = parseClassLabel(label, "", "");
        return [info.grade, info.major, info.classNo].join("|");
      });
      return keys.indexOf(r.classKey) === -1;
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
        sheet: candidates.length > 1
          ? candidates.map(function (c) { return c.sheet; }).join(" / ")
          : chosen.sheet,
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
