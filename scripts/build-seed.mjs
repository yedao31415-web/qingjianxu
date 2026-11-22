#!/usr/bin/env node
/**
 * 生成 public/assets/data/seed.json
 *
 * 两种模式：
 *   1) 不带参数：生成完全虚构的示例数据（仓库默认，可安全公开）；
 *   2) 带 xlsx / xls / csv 路径：解析本机总表并生成种子。
 *
 * ⚠️ 真实总表与由它生成的 seed.json 含个人课表信息，只可用于本机或私有部署，
 *    请勿提交到公开仓库。公开发布时保持默认的示例数据即可，使用者可在页面里
 *    「我的 → 导入考试总表」自行导入，数据只存在其本机浏览器。
 *
 * 用法：node scripts/build-seed.mjs [输入文件...]
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

const SOURCE_DIR = path.join(ROOT, "data/source");
const OUT = path.join(ROOT, "public/assets/data/seed.json");

const args = process.argv.slice(2).filter((a) => !a.startsWith("-"));

/* ------------------------------ 默认：示例数据 ------------------------------ */
if (!args.length) {
  const { execFileSync } = await import("node:child_process");
  execFileSync(process.execPath, [path.join(__dirname, "demo-seed.mjs"), OUT], { stdio: "inherit" });
  console.log("提示：公开发布请保留示例数据；真实课表请在页面内导入，不要提交到仓库。");
  process.exit(0);
}

/* ------------------------------ 可选：本地总表 ------------------------------ */
const files = args.map((a) => path.resolve(a));
for (const file of files) {
  if (!fs.existsSync(file)) {
    console.error(`源文件不存在：${file}`);
    process.exit(1);
  }
}

console.warn("⚠️ 你正在用真实总表生成种子数据：请勿把 data/source/ 与 seed.json 提交到公开仓库。");
console.warn("   公开部署请改用 `npm run build:seed`（默认示例数据），真实课表在页面内导入。\n");

const datasets = [];
for (const file of files) {
  const buf = fs.readFileSync(file);
  const wb = XLSX.read(buf, { type: "buffer", cellDates: false });
  const base = path.basename(file);
  const res = Parser.parseWorkbook(wb, XLSX, {
    sourceName: base,
    title: "",
    datasetId: "ds-" + base.replace(/\.[^.]+$/, "").replace(/[^a-zA-Z0-9_-]/g, "-").toLowerCase()
  });

  if (!res.ok) {
    console.error(`[skip] ${base}: ${res.error}`);
    continue;
  }

  const ds = res.dataset;
  datasets.push({
    meta: {
      id: ds.id,
      title: ds.title,
      campus: ds.campus || "",
      term: ds.term || "",
      sourceName: ds.sourceName,
      sheet: ds.sheet,
      headerRow: ds.headerRow,
      summary: ds.summary,
      ownerId: "shared",
      builtin: true,
      importedAt: ds.importedAt
    },
    records: res.records.map((r) => ({
      id: r.id,
      course: r.course,
      code: r.code,
      college: r.college,
      date: r.date,
      start: r.start,
      end: r.end,
      weekday: r.weekday,
      week: r.week,
      grade: r.grade,
      major: r.major,
      classNo: r.classNo,
      cls: r.cls,
      classKey: r.classKey,
      headcount: r.headcount,
      rooms: r.rooms,
      teaching: r.teaching,
      iso: r.iso
    }))
  });

  console.log(
    `[ok] ${base}\n     课程 ${ds.summary.courseCount} · 记录 ${ds.summary.recordCount} · 班级 ${ds.summary.classCount} · ${ds.summary.dateFrom} ~ ${ds.summary.dateTo}`
  );
  res.warnings.forEach((w) => console.log(`      - ${w}`));
}

if (!datasets.length) {
  console.error("没有可用的总表，已保留原有 seed.json。");
  process.exit(1);
}

const seed = { version: 1, generatedAt: new Date().toISOString(), datasets };

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, JSON.stringify(seed), "utf8");
const kb = (fs.statSync(OUT).size / 1024).toFixed(1);
console.log(`\n✓ 写入 ${path.relative(ROOT, OUT)}（${kb} KB, ${datasets.length} 个总表）`);
