#!/usr/bin/env node
/**
 * 发布目录白名单校验。
 *
 * 规则：public/ 只允许「运行必需文件」存在。
 * 任何源数据（xlsx / xls / csv）、README、skills、docs、data、脚本、工具目录都会让校验失败。
 * 用于 CI 与本地 npm run check，避免把源表或开发文件带上 GitHub Pages。
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const PUBLIC = path.join(ROOT, "public");

const REQUIRED = [
  "index.html",
  "manifest.webmanifest",
  "sw.js",
  ".nojekyll",
  "assets/icon.svg",
  "assets/icon-192.png",
  "assets/icon-512.png",
  "assets/icon-maskable-512.png",
  "assets/credit-sign.png",
  "assets/css/app.css",
  "assets/js/app.js",
  "assets/js/db.js",
  "assets/js/parser.js",
  "assets/js/export.js",
  "assets/data/seed.json",
  "vendor/xlsx.full.min.js",
  "vendor/jspdf.umd.min.js",
  "vendor/html2canvas.min.js",
];

/**
 * 允许的文件：精确路径或规则。
 * 仅这些扩展名可发布；其余一律拒绝。
 */
const ALLOWED_EXT = new Set([".html", ".css", ".js", ".json", ".png", ".svg", ".webmanifest"]);
const FORBIDDEN_NAME = /(^|\/)(readme|skills?|docs?|data\/source|scripts?|tools?)(\/|\.|$)/i;
const FORBIDDEN_EXT = new Set([".xlsx", ".xls", ".csv", ".md", ".mjs", ".map", ".zip", ".psd", ".ts"]);

function walk(dir, acc = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full, acc);
    else acc.push(full);
  }
  return acc;
}

if (!fs.existsSync(PUBLIC)) {
  console.error("✗ public/ 不存在");
  process.exit(1);
}

const files = walk(PUBLIC);
const errors = [];
let total = 0;

for (const full of files) {
  const rel = path.relative(PUBLIC, full).split(path.sep).join("/");
  const ext = path.extname(rel).toLowerCase();
  const stat = fs.statSync(full);
  total += stat.size;

  if (FORBIDDEN_NAME.test(rel)) errors.push(`禁止的目录/文件名：${rel}`);
  if (FORBIDDEN_EXT.has(ext)) errors.push(`禁止的扩展名（${ext}）：${rel}`);
  if (!ALLOWED_EXT.has(ext) && rel !== ".nojekyll") errors.push(`未在白名单中的文件类型：${rel}`);
  if (stat.size > 1.6 * 1024 * 1024) errors.push(`单文件超过 1.6 MB：${rel}`);
}

for (const rel of REQUIRED) {
  if (!fs.existsSync(path.join(PUBLIC, rel))) errors.push(`缺少运行必需文件：${rel}`);
}

// 明确拦截源数据上传，即使扩展名被改
for (const full of files) {
  const rel = path.relative(PUBLIC, full).split(path.sep).join("/");
  if (/(源表|课表|考试安排|成绩|名单|\.xlsx?$|\.csv$)/i.test(rel)) errors.push(`疑似源表进入发布目录：${rel}`);
}

// 公开仓库只允许示例数据：内置总表必须带 demo 标记
if (process.env.ALLOW_REAL_SEED !== "1") {
  try {
    const seed = JSON.parse(fs.readFileSync(path.join(PUBLIC, "assets/data/seed.json"), "utf8"));
    const real = (seed.datasets || []).filter((d) => !d.meta || d.meta.demo !== true);
    if (real.length) {
      errors.push(
        `seed.json 含非示例总表（${real.map((d) => (d.meta && d.meta.title) || "未命名").join("、")}）；` +
          "公开部署请运行 npm run seed 恢复示例数据，确需私有部署可加 ALLOW_REAL_SEED=1"
      );
    }
  } catch (e) {
    errors.push("seed.json 无法解析：" + e.message);
  }
}

if (errors.length) {
  console.error("✗ 发布目录校验失败：\n");
  errors.forEach((e) => console.error("  - " + e));
  process.exit(1);
}

console.log(`✓ 发布目录校验通过：${files.length} 个文件，${(total / 1024 / 1024).toFixed(2)} MB`);
console.log("  仅含运行必需文件；无源表、README、skills、docs、脚本。");
