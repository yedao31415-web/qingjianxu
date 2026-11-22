#!/usr/bin/env node
/**
 * 零依赖静态服务器：node scripts/dev-server.mjs [port] [root]
 * root 默认 public（发布目录），用于验证相对路径与子路径行为。
 */
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", process.argv[3] || "public");
const PORT = Number(process.argv[2] || process.env.PORT || 4176);

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".webmanifest": "application/manifest+json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
};

http
  .createServer((req, res) => {
    let urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
    if (urlPath === "/") urlPath = "/index.html";
    else if (urlPath.endsWith("/")) urlPath += "index.html";
    const filePath = path.join(ROOT, urlPath);
    if (!filePath.startsWith(ROOT)) { res.writeHead(403).end("Forbidden"); return; }

    fs.stat(filePath, (err, stat) => {
      if (err || !stat.isFile()) {
        res.writeHead(404, { "content-type": "text/plain; charset=utf-8" }).end("404 Not Found");
        return;
      }
      const ext = path.extname(filePath).toLowerCase();
      res.writeHead(200, {
        "content-type": TYPES[ext] || "application/octet-stream",
        "cache-control": "no-cache"
      });
      fs.createReadStream(filePath).pipe(res);
    });
  })
  .listen(PORT, "127.0.0.1", () => {
    console.log(`青简叙 dev server → http://127.0.0.1:${PORT}`);
  });
