#!/usr/bin/env node
/**
 * 青简叙 Find —— 命令行查考试（find skill 的可执行部分）
 *
 * 数据源可直接读取 public/assets/data/seed.json，也可读取用户导出的
 * 「青简叙数据包」(*.json)，因此可以脱离浏览器在终端 / Agent 中检索。
 *
 * 用法示例：
 *   node skills/exam-finder/find.mjs --class "24级汉语言01班"
 *   node skills/exam-finder/find.mjs --grade 26级 --major 计科 --upcoming
 *   node skills/exam-finder/find.mjs --course 数据结构 --json
 *   node skills/exam-finder/find.mjs --class "25级法学01班" --dataset 0
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

/* ------------------------------ 参数解析 ------------------------------ */
const argv = process.argv.slice(2);
const arg = (name, fallback = "") => {
  const i = argv.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const v = argv[i + 1];
  return v && !v.startsWith("--") ? v : "true";
};
const has = (name) => argv.includes(`--${name}`);

const DATA = path.resolve(ROOT, arg("data", "public/assets/data/seed.json"));
const CLASS = arg("class");
const GRADE = arg("grade");
const MAJOR = arg("major");
const COURSE = arg("course");
const DATE = arg("date");
const DATASET = arg("dataset", "");
const UPCOMING = has("upcoming");
const JSON_OUT = has("json");
const LIMIT = Number(arg("limit", "0")) || 0;

if (has("help") || argv.length === 0 && false) {
  console.log("用法: node skills/exam-finder/find.mjs --class <班级> [--grade] [--major] [--course] [--date] [--upcoming] [--json]");
  process.exit(0);
}

/* ------------------------------ 载入数据 ------------------------------ */
if (!fs.existsSync(DATA)) {
  console.error(`数据文件不存在: ${DATA}\n请先运行 npm run build:seed 生成，或指定 --data <文件>`);
  process.exit(2);
}

const raw = JSON.parse(fs.readFileSync(DATA, "utf8"));
const datasets = Array.isArray(raw) ? [{ meta: { title: "未命名" }, records: raw }] : raw.datasets || [];

const rows = [];
datasets.forEach((ds, i) => {
  if (DATASET && String(i) !== DATASET && !(ds.meta.title || "").includes(DATASET)) return;
  (ds.records || []).forEach((r) => rows.push(Object.assign({ __dataset: ds.meta.title || `总表${i + 1}` }, r)));
});

/* ------------------------------ 过滤 ------------------------------ */
const norm = (s) => String(s || "").replace(/\s+/g, "").toLowerCase();
const now = Date.now();

let hits = rows.filter((r) => {
  if (CLASS && norm(r.cls) !== norm(CLASS)) return false;
  if (GRADE && norm(r.grade) !== norm(GRADE)) return false;
  if (MAJOR && norm(r.major) !== norm(MAJOR)) return false;
  if (COURSE && !norm(r.course).includes(norm(COURSE))) return false;
  if (DATE && String(r.date) !== DATE) return false;
  if (UPCOMING && Date.parse(r.iso || `${r.date}T23:59:00`) < now) return false;
  return true;
});

hits.sort((a, b) => `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`) || a.course.localeCompare(b.course, "zh"));

if (LIMIT) hits = hits.slice(0, LIMIT);

/* ------------------------------ 输出 ------------------------------ */
const WEEK = ["日", "一", "二", "三", "四", "五", "六"];
const weekday = (d) => {
  const t = new Date(`${d}T00:00:00`);
  return isNaN(t) ? "" : `周${WEEK[t.getDay()]}`;
};

if (JSON_OUT) {
  console.log(JSON.stringify({ count: hits.length, records: hits }, null, 2));
  process.exit(hits.length ? 0 : 1);
}

if (!hits.length) {
  console.error("未找到匹配的考试记录。可用 --help 查看参数。");
  process.exit(1);
}

const width = (s) => Array.from(String(s)).reduce((n, ch) => n + (ch.charCodeAt(0) > 255 ? 2 : 1), 0);
const pad = (s, n) => String(s) + " ".repeat(Math.max(0, n - width(s)));

const table = hits.map((r) => [
  r.date,
  weekday(r.date),
  `${r.start || "--:--"}-${r.end || "--:--"}`,
  r.course,
  (r.rooms || []).join("、") || "—",
  r.cls || "多班合考",
  r.headcount == null ? "" : String(r.headcount)
]);

const headers = ["日期", "星期", "时间", "课程", "地点", "班级", "人数"];
const widths = headers.map((h, i) => Math.max(width(h), ...table.map((row) => width(row[i]))) + 2);

console.log(headers.map((h, i) => pad(h, widths[i])).join(""));
console.log(widths.map((w) => "-".repeat(w - 1)).join(" "));
table.forEach((row) => console.log(row.map((c, i) => pad(c, widths[i])).join("")));
console.log(`\n共 ${hits.length} 条 · 数据源 ${path.basename(DATA)} · 生成于 ${new Date().toLocaleString("zh-CN", { hour12: false })}`);
