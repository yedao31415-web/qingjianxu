# 源数据说明（仅本机，不随仓库发布）

本目录**不包含任何真实课表**。仓库默认的 `public/assets/data/seed.json` 由
`scripts/demo-seed.mjs` 生成，学校、院系、专业、班级、课程、教室、人数全部为虚构示例。

## 生成示例数据（默认，可安全公开）

```bash
npm run seed
```

- 日期以运行时的「本周周一」为锚点，向后排三周，保证随时都有可倒数的场次；
- 生成结果 162 条记录、39 门课程、18 个班级，覆盖 24 / 25 / 26 级五个专业；
- 输出到 `public/assets/data/seed.json`，可随站点公开发布。

## 使用自己的真实总表（只在本机）

1. 把总表放在本目录（`*.xlsx` / `*.xls` / `*.csv` 已被 `.gitignore` 排除）；
2. 运行 `node scripts/build-seed.mjs data/source/你的总表.xlsx`；
3. **不要提交** 该总表与真实数据版本的 `seed.json`，也不要把它部署到公开站点；
4. 需要撤销时重新运行 `npm run seed` 恢复示例数据。

更推荐的做法：直接部署示例版本，打开网页后用「我的 → 导入考试总表」导入真实总表。
文件只在浏览器内解析，记录存进本机 IndexedDB，不会上传，也不进入仓库。

## 发布目录约束

源表不会复制到 `public/`，站点上不存在任何 xlsx 文件或下载链接；
`scripts/check-release.mjs` 会拒绝任何 xlsx / xls / csv 进入发布目录。
