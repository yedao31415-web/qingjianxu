/**
 * 生成完全虚构的示例考试数据（seed.json）。
 *
 * 设计原则：
 * 1. 学校、校区、院系、专业、班级、课程、教室、人数全部为虚构，不对应任何真实院校；
 * 2. 日期以「本周周一」为锚点向后排三周，保证任何时候生成都有可倒数的场次；
 * 3. 同一场考试（同年级同专业同课程）的所有班级共用日期、时段与教室，贴近真实排考；
 * 4. 输出结构与教务总表解析结果完全一致，前端无需任何特判。
 *
 * 用法：node scripts/demo-seed.mjs [输出路径]
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const OUT = path.resolve(process.argv[2] || path.join(ROOT, "public/assets/data/seed.json"));

/* ------------------------------ 虚构的基础设置 ------------------------------ */
const DATASET_ID = "ds-demo-2026-2027-1";
const CAMPUS = "云岫校区";
const TITLE = "2026-2027 学年第一学期 · 期末考试（示例数据）";
const TERM = "2026-2027-1";

const GRADES = ["24级", "25级", "26级"];

const MAJORS = [
  {
    name: "汉语言",
    code: "HY",
    college: "文学院~中国语言文学系",
    classes: { "24级": 2, "25级": 1, "26级": 1 },
    courses: [
      ["中国古代文学（三）", "文学史"],
      ["中国现当代文学", "文学史"],
      ["现代汉语（下）", "语言学"],
      ["文学理论导论", "文艺学"],
      ["外国文学史", "文学史"],
      ["写作学基础", "写作教研"]
    ]
  },
  {
    name: "计科",
    code: "JK",
    college: "计算机学院~计算机科学与技术系",
    classes: { "24级": 2, "25级": 2, "26级": 1 },
    courses: [
      ["数据结构与算法", "软件基础"],
      ["计算机组成原理", "硬件基础"],
      ["操作系统", "系统软件"],
      ["计算机网络", "网络工程"],
      ["数据库系统", "数据科学"],
      ["高等数学（上）", "数学公共课"]
    ]
  },
  {
    name: "法学",
    code: "FX",
    college: "法学院~民商法系",
    classes: { "24级": 1, "25级": 1, "26级": 1 },
    courses: [
      ["民法学（二）", "民商法"],
      ["刑法总论", "刑法学"],
      ["宪法学", "理论法学"],
      ["行政法与行政诉讼法", "宪法行政法"],
      ["国际法", "国际法学"],
      ["法律职业伦理", "理论法学"]
    ]
  },
  {
    name: "会计",
    code: "KJ",
    college: "商学院~会计学系",
    classes: { "24级": 1, "25级": 1, "26级": 1 },
    courses: [
      ["中级财务会计", "会计学"],
      ["成本会计", "会计学"],
      ["审计学基础", "审计学"],
      ["财务管理", "财务学"],
      ["税法", "财税法学"],
      ["管理会计", "会计学"]
    ]
  },
  {
    name: "药学",
    code: "YX",
    college: "药学院~药学系",
    classes: { "24级": 1, "25级": 1, "26级": 1 },
    courses: [
      ["药理学", "药理毒理"],
      ["药物化学", "药物化学"],
      ["药剂学", "药剂学"],
      ["药物分析", "药物分析"],
      ["生药学", "生药学"],
      ["药事管理学", "药事管理"]
    ]
  }
];

const COMMON_COURSES = {
  "24级": [
    ["大学英语（三）", "外国语学院~大学英语部"],
    ["马克思主义基本原理", "马克思主义学院"],
    ["体育（三）", "体育部"]
  ],
  "25级": [
    ["大学英语（二）", "外国语学院~大学英语部"],
    ["中国近现代史纲要", "马克思主义学院"],
    ["体育（二）", "体育部"]
  ],
  "26级": [
    ["大学英语（一）", "外国语学院~大学英语部"],
    ["思想道德与法治", "马克思主义学院"],
    ["体育（一）", "体育部"]
  ]
};

const ROOMS = [
  "博雅楼 A201",
  "博雅楼 A305",
  "明理楼 B102",
  "明理楼 B208",
  "知行楼 C301",
  "博学楼 D105"
];

const SLOTS = [
  ["09:00", "11:00"],
  ["14:30", "16:30"],
  ["19:00", "21:00"]
];

/* ------------------------------ 日期工具 ------------------------------ */
const pad = (n) => String(n).padStart(2, "0");
const fmt = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const WEEKDAY_CN = ["日", "一", "二", "三", "四", "五", "六"];

function mondayOfWeek(base) {
  const d = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  const offset = (d.getDay() + 6) % 7; // 周一为 0
  d.setDate(d.getDate() - offset);
  return d;
}

function nextWeekdays(start, count) {
  const out = [];
  const d = new Date(start.getTime());
  while (out.length < count) {
    const wd = d.getDay();
    if (wd !== 0 && wd !== 6) out.push(new Date(d.getTime()));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

/* ------------------------------ 构建记录 ------------------------------ */
const weekdays = nextWeekdays(mondayOfWeek(new Date()), 15);

const records = [];
let seq = 0;

function pushRecord(ctx) {
  const { course, code, college, dayIndex, slot, grade, major, classList } = ctx;
  const date = weekdays[dayIndex];
  const [start, end] = SLOTS[slot];
  const dateStr = fmt(date);
  const room = ROOMS[(dayIndex + slot) % ROOMS.length];
  const headcount = 28 + (classList.length * 2);
  classList.forEach((cls, idx) => {
    const classNo = cls.replace(/班$/, "").slice(-2);
    records.push({
      id: `${DATASET_ID}#${++seq}`,
      course,
      code,
      college,
      date: dateStr,
      start,
      end,
      weekday: WEEKDAY_CN[date.getDay()],
      week: 11 + Math.floor(dayIndex / 5),
      grade,
      major,
      classNo,
      cls,
      classKey: `${grade}|${major}|${classNo}`,
      headcount: headcount - (idx % 2),
      rooms: [room],
      teaching: classList.join(";"),
      iso: `${dateStr}T${end}:00`
    });
  });
}

GRADES.forEach((grade, gi) => {
  // 1) 专业必修课：同年级同专业的所有班级合考
  MAJORS.forEach((major, mi) => {
    const count = major.classes[grade] || 0;
    if (!count) return;
    const classList = Array.from({ length: count }, (_, i) => `${grade}${major.name}${pad(i + 1)}班`);
    major.courses.forEach(([course], ci) => {
      const dayIndex = (ci * 2 + mi * 3) % 11;
      pushRecord({
        course,
        code: `QJ-${major.code}${grade.slice(0, 2)}${pad(ci + 1)}`,
        college: major.college,
        dayIndex,
        slot: ci % 3,
        grade,
        major: major.name,
        classList
      });
    });
  });

  // 2）公共课：同年级跨专业统考，排在最后三天，避开专业课
  const gradeMajors = MAJORS.filter((m) => m.classes[grade]);
  COMMON_COURSES[grade].forEach(([course, college], ci) => {
    gradeMajors.forEach((m) => {
      const classList = Array.from(
        { length: m.classes[grade] },
        (_, i) => `${grade}${m.name}${pad(i + 1)}班`
      );
      pushRecord({
        course,
        code: `QJ-GG${grade.slice(0, 2)}${pad(ci + 1)}`,
        college,
        dayIndex: 11 + ((gi + ci) % 3),
        slot: (gi + ci) % 3,
        grade,
        major: m.name,
        classList
      });
    });
  });
});

records.sort(
  (a, b) => (a.date + a.start).localeCompare(b.date + b.start) || a.cls.localeCompare(b.cls, "zh")
);

const grades = [...new Set(records.map((r) => r.grade))].sort();
const majors = [...new Set(records.map((r) => r.major).filter(Boolean))];
const classes = [...new Set(records.map((r) => r.cls))].sort((a, b) => a.localeCompare(b, "zh"));
const courseCount = new Set(records.map((r) => r.course)).size;

const seed = {
  version: 1,
  generatedAt: new Date().toISOString(),
  datasets: [
    {
      meta: {
        id: DATASET_ID,
        title: TITLE,
        campus: CAMPUS,
        term: TERM,
        sourceName: "",
        sheet: "",
        headerRow: 0,
        ownerId: "shared",
        builtin: true,
        demo: true,
        importedAt: new Date().toISOString(),
        summary: {
          recordCount: records.length,
          courseCount,
          classCount: classes.length,
          grades,
          majors,
          classes,
          dateFrom: records[0] ? records[0].date : "",
          dateTo: records.length ? records[records.length - 1].date : ""
        }
      },
      records
    }
  ]
};

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, JSON.stringify(seed), "utf8");

const kb = (fs.statSync(OUT).size / 1024).toFixed(1);
console.log(
  `[demo] 示例数据已生成：${records.length} 条记录 · ${courseCount} 门课程 · ${classes.length} 个班级 · ` +
    `${seed.datasets[0].meta.summary.dateFrom} ~ ${seed.datasets[0].meta.summary.dateTo}`
);
console.log(`[demo] 写入 ${path.relative(ROOT, OUT)}（${kb} KB）`);
