#!/usr/bin/env node
/**
 * 考试总表解析兼容性测试。
 *
 * 用法：
 *   node scripts/test-parser.mjs                  # 跑内置的 14 种表格格式
 *   node scripts/test-parser.mjs 文件1.xlsx ...   # 解析自己的总表并输出诊断
 *
 * 内置样例全部为虚构数据，不包含任何真实课表。
 */
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const require = createRequire(import.meta.url);
const XLSX = require(path.join(ROOT, "public/vendor/xlsx.full.min.js"));
const Parser = require(path.join(ROOT, "public/assets/js/parser.js"));

/* ------------------------------ 工具 ------------------------------ */
let pass = 0;
let fail = 0;
const failures = [];

function check(name, cond, detail) {
  if (cond) {
    pass++;
    return true;
  }
  fail++;
  failures.push(`${name}${detail ? "：" + detail : ""}`);
  return false;
}

function excelSerial(y, m, d, hh = 0, mm = 0) {
  return (Date.UTC(y, m - 1, d, hh, mm) - Date.UTC(1899, 11, 30)) / 86400000;
}

function makeBook(sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach((sheet, i) => {
    // 支持两种写法：[名称, 行数组, 合并区域] 或直接传行数组
    const named = typeof sheet[0] === "string";
    const name = named ? sheet[0] : "Sheet" + (i + 1);
    const rows = named ? sheet[1] : sheet;
    const merges = named ? sheet[2] : null;
    const ws = XLSX.utils.aoa_to_sheet(rows);
    if (merges && merges.length) ws["!merges"] = merges;
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  return wb;
}

function parse(wb) {
  return Parser.parseWorkbook(wb, XLSX, { datasetId: "ds-test", sourceName: "fixture.xlsx" });
}

const HEADER = ["序号", "课程名称", "课程代码", "考试日期", "考试时间", "考场", "班级", "人数"];

function row(course, date, time, room, cls, count = 30, code = "C001") {
  return [1, course, code, date, time, room, cls, count];
}

/* ------------------------------ 用例 ------------------------------ */
const CASES = [];
function test(name, fn) {
  CASES.push({ name, fn });
}

test("标准中文表头 + 文本日期", () => {
  const res = parse(
    makeBook([
      [
        ["2026-2027 学年第一学期期末考试安排"],
        HEADER,
        row("高等数学（上）", "2027-01-05", "09:00-11:00", "明理楼A101", "2024级软件工程1班", 40),
        row("大学英语（三）", "2027-01-07", "14:30-16:30", "博雅楼B201", "2024级软件工程1班", 40),
        row("数据结构", "2027-01-09", "19:00-21:00", "知行楼C301", "2024级计算机科学与技术（1）班", 38)
      ]
    ])
  );
  check("标准表：可解析", res.ok);
  if (!res.ok) return;
  check("标准表：记录数", res.dataset.summary.recordCount === 3, res.dataset.summary.recordCount);
  const r = res.records[0];
  check("标准表：日期", r.date === "2027-01-05", r.date);
  check("标准表：时段", r.start === "09:00" && r.end === "11:00", r.start + "-" + r.end);
  check("标准表：年级", r.grade === "24级", r.grade);
  check("标准表：专业", r.major === "软件工程", r.major);
  check("标准表：班号", r.classNo === "01", r.classNo);
  check("标准表：班级", r.cls === "24级软件工程01班", r.cls);
  check("标准表：表头行", res.dataset.headerRow === 2, res.dataset.headerRow);
  check("标准表：括号班号", res.records[2].classNo === "01", res.records[2].classNo);
});

test("Excel 日期序列号 + 时间文本", () => {
  const serial = excelSerial(2027, 1, 11);
  const res = parse(
    makeBook([
      [
        ["课程名称", "考试日期", "考试时间", "考试地点", "行政班"],
        ["线性代数", serial, "09:00-11:00", "A101", "2024级数学1班"],
        ["概率论", serial + 2, "14:30-16:30", "A102", "2024级数学1班"]
      ]
    ])
  );
  check("序列号：可解析", res.ok);
  if (!res.ok) return;
  check("序列号：记录数", res.dataset.summary.recordCount === 2);
  check("序列号：日期", res.records[0].date === "2027-01-11", res.records[0].date);
  check("序列号：第二场", res.records[1].date === "2027-01-13", res.records[1].date);
});

test("日期 / 开始时间 / 结束时间 分三列", () => {
  const res = parse(
    makeBook([
      [
        ["考试科目", "考试日期", "开始时间", "结束时间", "考试教室", "学生班级"],
        ["大学物理", "2027年1月12日", "09:00", "11:00", "综合楼201", "24级物理1班"],
        ["量子力学", "2027年1月14日", "14:30", "16:30", "综合楼202", "24级物理1班"]
      ]
    ])
  );
  check("分列时间：可解析", res.ok);
  if (!res.ok) return;
  check("分列时间：起止", res.records[0].start === "09:00" && res.records[0].end === "11:00",
    res.records[0].start + "-" + res.records[0].end);
  check("分列时间：中文日期", res.records[1].date === "2027-01-14", res.records[1].date);
});

test("月日无年份（跨年推断）", () => {
  const res = parse(
    makeBook([
      [
        ["课程", "日期", "时间", "地点", "班级"],
        ["中国近现代史", "1月5日", "上午 9:00—11:00", "B101", "24级会计1班"]
      ]
    ])
  );
  check("无年份：可解析", res.ok);
  if (!res.ok) return;
  const r = res.records[0];
  check("无年份：日期含年份", /^\d{4}-01-05$/.test(r.date), r.date);
  check("无年份：未落到过去", +r.date.slice(0, 4) >= new Date().getFullYear(), r.date);
  check("无年份：上午时段", r.start === "09:00" && r.end === "11:00", r.start + "-" + r.end);
});

test("班级写法：括号 / 四位数 / 短横线 / 四位年份", () => {
  const res = parse(
    makeBook([
      [
        ["课程名称", "考试日期", "考试时间", "考场", "班级"],
        ["甲", "2027-01-05", "09:00-11:00", "A101", "2024级计算机科学与技术（1）班"],
        ["乙", "2027-01-06", "09:00-11:00", "A102", "计科2402班"],
        ["丙", "2027-01-07", "09:00-11:00", "A103", "软件24-3班"],
        ["丁", "2027-01-08", "09:00-11:00", "A104", "2024法学1班"]
      ]
    ])
  );
  check("班级写法：可解析", res.ok);
  if (!res.ok) return;
  const got = res.records.map((r) => [r.grade, r.major, r.classNo].join("/"));
  check("班级写法：括号", got[0] === "24级/计算机科学与技术/01", got[0]);
  check("班级写法：计科2402", got[1] === "24级/计科/02", got[1]);
  check("班级写法：软件24-3", got[2] === "24级/软件/03", got[2]);
  check("班级写法：2024法学1", got[3] === "24级/法学/01", got[3]);
});

test("只有「考试班级」且多班合考", () => {
  const res = parse(
    makeBook([
      [
        ["课程名称", "考试日期", "考试时间", "考试地点", "考试班级", "人数"],
        ["大学物理", "2027-01-15", "09:00-11:00", "博雅楼A201", "计科2401班、计科2402班、软件2401班", 90]
      ]
    ])
  );
  check("合考拆行：可解析", res.ok);
  if (!res.ok) return;
  check("合考拆行：记录数", res.dataset.summary.recordCount === 3, res.dataset.summary.recordCount);
  check("合考拆行：班级数", res.dataset.summary.classCount === 3, res.dataset.summary.classCount);
  check("合考拆行：班号", res.records.map((r) => r.classNo).sort().join(",") === "01,01,02",
    res.records.map((r) => r.cls).join(","));
  check("合考拆行：人数均分", res.records.every((r) => r.headcount === 30), res.records.map((r) => r.headcount).join(","));
});

test("多工作表合并 + 重复行去重", () => {
  const h = ["课程名称", "考试日期", "考试时间", "考场", "班级"];
  const a = [h.slice(), ["高数", "2027-01-05", "09:00-11:00", "A101", "24级计科1班"]];
  const b = [h.slice(), ["英语", "2027-01-06", "09:00-11:00", "A102", "24级计科1班"]];
  const c = [h.slice(), ["物理", "2027-01-07", "09:00-11:00", "A103", "24级计科1班"]];
  const res = parse(
    makeBook([
      ["学院一", a],
      ["学院二", b],
      ["学院三", c],
      ["学院一副本", a]
    ])
  );
  check("多表：可解析", res.ok);
  if (!res.ok) return;
  check("多表：合并记录数", res.dataset.summary.recordCount === 3, res.dataset.summary.recordCount);
  check("多表：sheet 标注", res.dataset.sheet.includes(" / "), res.dataset.sheet);
  check("多表：重复提示", res.warnings.some((w) => w.includes("忽略重复")), res.warnings.join(" | "));
});

test("合并单元格：课程/日期纵向覆盖多个班", () => {
  const rows = [
    ["课程名称", "考试日期", "考试时间", "考场", "班级"],
    ["概率论", "2027-01-08", "09:00-11:00", "A201", "24级数学1班"],
    [null, null, null, null, "24级数学2班"]
  ];
  const merges = [0, 1, 2, 3].map((c) => ({ s: { r: 1, c }, e: { r: 2, c } }));
  const res = parse(makeBook([["合并表", rows, merges]]));
  check("合并单元格：可解析", res.ok);
  if (!res.ok) return;
  check("合并单元格：记录数", res.dataset.summary.recordCount === 2, res.dataset.summary.recordCount);
  check("合并单元格：课程", res.records.every((r) => r.course === "概率论"));
  check("合并单元格：日期", res.records.every((r) => r.date === "2027-01-08"));
  check("合并单元格：班级", res.records.map((r) => r.classNo).sort().join(",") === "01,02");
});

test("表头位于较深位置", () => {
  const rows = [
    ["XX大学考试安排"],
    ["说明：请携带学生证"],
    [],
    ["制表：教务处"],
    ["2026-12-01 发布"],
    ["课程名称", "考试日期", "考试时间", "考场", "班级"],
    ["形势与政策", "2027-01-09", "14:30-16:30", "A301", "24级法学1班"]
  ];
  const res = parse(makeBook([["说明", rows]]));
  check("深表头：可解析", res.ok);
  if (!res.ok) return;
  check("深表头：表头行", res.dataset.headerRow === 6, res.dataset.headerRow);
  check("深表头：记录数", res.dataset.summary.recordCount === 1);
});

test("只有年级/专业、没有班级", () => {
  const res = parse(
    makeBook([
      [
        ["课程名称", "考试日期", "考试时间", "考场", "年级", "专业"],
        ["有机化学", "2027-01-18", "09:00-11:00", "D105", "2024级", "药学"]
      ]
    ])
  );
  check("无班级：可解析", res.ok);
  if (!res.ok) return;
  check("无班级：年级专业", res.records[0].grade === "24级" && res.records[0].major === "药学",
    res.records[0].grade + "/" + res.records[0].major);
  check("无班级：提示", res.warnings.some((w) => w.includes("没有具体班级")), res.warnings.join(" | "));
});

test("英文表头", () => {
  const res = parse(
    makeBook([
      [
        ["Course Name", "Exam Date", "Start Time", "End Time", "Venue", "Class"],
        ["Calculus", "2027-01-19", "09:00", "11:00", "Room A101", "2024软件工程1班"]
      ]
    ])
  );
  check("英文表头：可解析", res.ok);
  if (!res.ok) return;
  check("英文表头：时段", res.records[0].start === "09:00" && res.records[0].end === "11:00",
    res.records[0].start + "-" + res.records[0].end);
});

test("CSV 文件", () => {
  const csv = [
    "课程名称,考试日期,考试时间,考场,班级",
    "离散数学,2027-01-20,09:00-11:00,A401,24级计科1班",
    "编译原理,2027-01-21,14:30-16:30,A402,24级计科1班"
  ].join("\n");
  const res = parse(XLSX.read(csv, { type: "string" }));
  check("CSV：可解析", res.ok);
  if (!res.ok) return;
  check("CSV：记录数", res.dataset.summary.recordCount === 2);
});

test("列顺序打乱", () => {
  const res = parse(
    makeBook([
      [
        ["班级", "考场", "考试时间", "课程名称", "考试日期"],
        ["24级计科1班", "A501", "09:00-11:00", "人工智能导论", "2027-01-22"]
      ]
    ])
  );
  check("乱序：可解析", res.ok);
  if (!res.ok) return;
  check("乱序：字段", res.records[0].course === "人工智能导论" && res.records[0].rooms[0] === "A501");
});

test("直接测试班级与时间解析函数", () => {
  const cases = [
    ["2024级计算机科学与技术（1）班", "24级", "计算机科学与技术", "01"],
    ["计科2402班", "24级", "计科", "02"],
    ["软件24-3班", "24级", "软件", "03"],
    ["2024法学1班", "24级", "法学", "01"],
    ["23级湛临床05班", "23级", "湛临床", "05"],
    ["2401班", "24级", "", "01"]
  ];
  cases.forEach(([input, g, m, c], i) => {
    const r = Parser.parseClassLabel(input, "", "");
    check("班级函数 #" + (i + 1), r.grade === g && r.major === m && r.classNo === c,
      `${input} → ${r.grade}/${r.major}/${r.classNo}`);
  });

  const t1 = Parser.parseExamTime("09:00-11:00", excelSerial(2027, 2, 3), 2026);
  check("时间函数：序列号日期", t1.date === "2027-02-03", t1.date);

  const t2 = Parser.parseExamTime("", "2027-02-04 14:30-16:30", 2026);
  check("时间函数：日期列带时段", t2.date === "2027-02-04" && t2.start === "14:30" && t2.end === "16:30",
    `${t2.date} ${t2.start}-${t2.end}`);
});

/* ------------------------------ 执行 ------------------------------ */
if (process.argv.length > 2) {
  for (const file of process.argv.slice(2)) {
    const buf = fs.readFileSync(path.resolve(file));
    const wb = XLSX.read(buf, { type: "buffer" });
    const res = Parser.parseWorkbook(wb, XLSX, { sourceName: path.basename(file) });
    console.log("\n=== " + path.basename(file) + " ===");
    if (!res.ok) {
      console.log("✗ " + res.error);
      process.exitCode = 1;
      continue;
    }
    const s = res.dataset.summary;
    console.log(`✓ ${s.recordCount} 条 · ${s.courseCount} 门课 · ${s.classCount} 个班 · ${s.dateFrom} ~ ${s.dateTo}`);
    console.log("表头行：" + res.dataset.headerRow + " · 工作表：" + res.dataset.sheet);
    console.log("识别列：" + Object.entries(res.dataset.columns).map(([k, v]) => k + "=" + v).join(", "));
    console.log("年级：" + s.grades.join("、"));
    console.log("专业：" + s.majors.join("、"));
    console.log("班级示例：" + s.classes.slice(0, 6).join("、"));
    res.warnings.forEach((w) => console.log("  · " + w));
    console.log("首条：" + JSON.stringify(res.records[0]));
  }
  process.exit(process.exitCode || 0);
}

console.log("考试总表解析兼容性测试\n");
for (const c of CASES) {
  const before = fail;
  c.fn();
  console.log((fail === before ? "✓ " : "✗ ") + c.name);
}
console.log(`\n通过 ${pass} 项检查，失败 ${fail} 项。`);
if (fail) {
  console.log("\n失败详情：");
  failures.forEach((f) => console.log("  - " + f));
  process.exit(1);
}
