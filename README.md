# 青简叙 · QingJianXu

> 以青简记考程，叙一季光阴。

**在线使用**：<https://yedao31415-web.github.io/qingjianxu/>

把教务下发的考试安排 Excel，变成手机上一页能查、能倒数、能自己补录实验与考试的「青简」。
纯静态、零运行时依赖、无构建步骤，数据只存本机浏览器。

<table>
  <tr>
    <td><img src="docs/screenshots/countdown.png" width="230" alt="倒数日" /></td>
    <td><img src="docs/screenshots/exams.png" width="230" alt="考程表" /></td>
    <td><img src="docs/screenshots/wallpaper.png" width="230" alt="自定义壁纸" /></td>
  </tr>
  <tr>
    <td align="center">倒数日 + 自建条目</td>
    <td align="center">筛选与检索</td>
    <td align="center">自定义壁纸</td>
  </tr>
</table>

## 这是什么

3.0 版本，独立命名为「青简叙」。它在保留 2.0 全部能力的基础上，新增了**手动录入实验 / 考试**与**倒数日自定义壁纸**，并把发布产物收敛为「只含运行必需文件」。

浏览器打开即用：导入总表 → 设置班级 → 看下一场考试还有几天 → 导出或分享。

## 与 2.0「考试雷达」的关系

3.0 是**独立目录、独立数据库、独立入口**，不会影响已发布 2.0 的任何功能或数据：

| 维度 | 2.0 考试雷达 | 3.0 青简叙 |
| --- | --- | --- |
| 目录 | `exam-radar/` | `qingjianxu/`（本目录） |
| 发布根 | 项目根目录 | `public/` |
| IndexedDB | `examradar` | `qingjianxu` |
| 缓存命名空间 | `examradar-*` | `qingjianxu-*` |
| 分享入口 | 原链接保持不变 | 新链接，单独发布 |
| 数据 | 各自浏览器内独立 | 各自浏览器内独立 |

两者不共享文件、数据库或缓存；打开 3.0 不会升级 2.0 的数据库，删除 3.0 的数据也不会影响 2.0。

## 功能

### 倒数日（首页）

- 只显示与使用者相关的考试，可在「本班 / 本专业 / 本年级 / 关注 / 全部」间切换。
- 大号天数 + 日期、时段、考场；已结束场次折叠，不显示负数天数。
- 每场考试按「课程 + 日期 + 起考时间」聚合，同一门课多班合考不会重复刷屏。
- **手动录入实验 / 考试**：名称、时间（日期 + 起止）、地点、备注；录入后立即并入同一条时间轴，最近的条目会成为首页倒数主角。
- **自定义壁纸**：导入任意图片，即时应用；可调模糊与遮罩，仅作用于倒数日页，按使用者档案保存。

### 考程表

- 年级 / 专业 / 班级三级联动筛选，班级列表支持关键词搜索。
- 搜索框同时匹配课程名、课程号、教室、班级、开课学院与自建条目备注。
- 结果按日期分组，首屏 120 条、按需加载，千行数据不卡顿。
- 星标关注课程，直接进入倒数日的「关注」范围。
- 自建条目与总表场次合并展示，用「实验 / 考试」朱砂标签区分，可直接点开编辑。

### 导出

- Excel（`.xlsx`，含考程表与说明页）、PDF（A4 逐页，中文免字体）、CSV（UTF-8 BOM）、Word（`.doc`，表头跨页重复）。
- 导出的永远是**当前筛选结果**；自建条目一并导出，「班级」列标注「自建实验 / 自建考试」，并新增「备注」列。

### 多总表与多使用者

- 可导入任意多份总表，顶部胶囊随时切换；查询只作用于当前总表。
- 每份总表有独立主键区间，删除只清除该表自身记录。
- 同机可建多个使用者档案，各自保存年级、班级、关注科目、自建条目与壁纸。
- 导入时可选「所有人共用」或「仅当前档案可见」。

### 分享与离线

- 分享链接自动带上班级（`#c=24级|汉语言|01`），同学打开即强制切到该班级视图，不受本机上次筛选影响。
- 「导出数据包」把整份总表发给同学，对方离线导入即可使用。
- Service Worker 缓存应用外壳，已打开过的页面离线仍可查询与倒数。

### 制作者标识

底部导航右端常驻「叶盗」书法签名与文字标识，三个页面均可见、不随滚动消失，也不遮挡列表内容；点击可打开「关于与数据说明」。

## 快速开始

无需 `npm install`（没有运行时依赖）。

```bash
npm run dev      # http://127.0.0.1:4176，直接预览发布目录
npm run check    # 校验 public/ 只含运行必需文件
npm run seed     # 重新生成脱敏示例数据
npm run find -- --class "24级汉语言01班"   # 命令行查考试
```

首次打开会自动载入内置示例数据（虚构的「示例大学 · 云岫校区」，162 条记录、18 个班级），仅用于演示界面与交互，**不含任何真实课表或个人信息**。

### 导入自己的总表

「我的 → 导入考试总表」，支持 `.xlsx` / `.xls` / `.csv`，可多选。

解析器会在前 16 行里寻找包含「课程名称 + 考试时间/日期 + 班级/年级/专业」的表头行，并识别常见列名别名。合并单元格按 `!merges` 真实区域回填，不做逐行向下填充，避免「一个时间合并覆盖多个课程」时串行错配。

导入后的记录只写入当前浏览器的 IndexedDB，**不会上传、不会写回仓库，也不会出现在 `seed.json`**。换设备时用「我的 → 导出数据包」迁移。

## 数据脱敏

仓库与发布包中的课表全部为虚构示例，可以安全公开：

| 项目 | 现状 |
| --- | --- |
| 内置数据 | `scripts/demo-seed.mjs` 生成，学校 / 院系 / 专业 / 班级 / 课程 / 教室 / 人数均为虚构 |
| 真实总表 | 不随仓库提供，`data/source/*.xlsx` 已被 `.gitignore` 排除 |
| 页面与文案 | 不含真实院校、班级、课程或个人信息 |
| 截图 | `docs/screenshots/` 全部由示例数据重新生成 |
| 运行时导入 | 用户自带的真实总表只留在其本机浏览器，不经服务器 |
| 部署校验 | `npm run check` 会拒绝非示例数据的 `seed.json`，避免误把真实课表推到公开站点 |

示例数据以运行时的「本周周一」为锚点排三周，任何时候重新生成都有可倒数的场次。`npm run seed` 即可重建。

> 确需把真实总表作为内置数据部署到**私有**仓库时，构建后使用 `ALLOW_REAL_SEED=1 npm run check` 跳过该校验。

## 部署到 GitHub Pages

### 一次性配置

1. 把本目录推送到 GitHub 仓库的 `main` 分支。
2. 仓库 **Settings → Pages → Build and deployment → Source** 选择 **GitHub Actions**。
3. 推送后 `.github/workflows/pages.yml` 自动执行：校验发布目录 → 上传 `public/` → 发布。

站点地址形如 `https://<用户名>.github.io/<仓库名>/`。

### 部署后再导入真实课表

公开站点只带示例数据。真实课表在**你自己的浏览器里**导入即可正常使用：

1. 打开线上地址，完成首次引导；
2. 「我的 → 导入考试总表」，选择教务下发的 Excel；
3. 在「总表管理」里切到刚导入的总表，回到「我的」设置年级 / 专业 / 班级；
4. 倒数日、考程表、四种导出、自建条目、壁纸全部照常工作。

数据留在该浏览器的 IndexedDB 中，刷新、离线都不丢；清除浏览器数据或换设备需要重新导入，也可先用「导出数据包」备份。

> 如果你希望真实课表成为站点内置数据（例如只给自己或小范围使用），可在本机运行 `node scripts/build-seed.mjs <总表路径>` 后部署到**私有仓库**。公开仓库请勿提交真实总表或真实版 `seed.json`。

### 子路径适配

GitHub Pages 项目站点部署在 `/仓库名/` 子路径下，本项目已按此约束处理：

- 所有资源引用均为相对路径（`assets/css/app.css`、`vendor/...`）；
- 数据加载使用相对 `fetch("assets/data/seed.json")`，随文档 URL 解析；
- `manifest.webmanifest` 的 `start_url` 为 `./index.html`，`scope` 为 `./`；
- Service Worker 以 `sw.js` 相对路径注册，作用域自动落在子路径内。

本地验证子路径行为：

```bash
mkdir -p /tmp/sub && cp -r public /tmp/sub/app
node scripts/dev-server.mjs 4177 /tmp
# 打开 http://127.0.0.1:4177/app/
```

### 发布目录约束

发布根是 `public/`，只允许运行必需文件：

```
public/
├── .nojekyll
├── index.html
├── manifest.webmanifest
├── sw.js
├── assets/
│   ├── icon.svg
│   ├── icon-192.png
│   ├── icon-512.png
│   ├── icon-maskable-512.png
│   ├── credit-sign.png
│   ├── css/app.css
│   ├── js/{app,db,parser,export}.js
│   └── data/seed.json
└── vendor/
    ├── xlsx.full.min.js
    ├── jspdf.umd.min.js
    └── html2canvas.min.js
```

**不会进入发布目录**：`data/source/*.xlsx`、`README.md`、`skills/`、`docs/`、`scripts/`、`.tools/`、`.github/`。`npm run check` 与 Pages 工作流会在部署前强制校验；任何源表、README 或开发文件混入都会让构建失败。

## 数据管线

```
scripts/demo-seed.mjs          （默认，虚构示例，可公开）
        │  npm run seed
        ▼
public/assets/data/seed.json   （示例数据，随站点发布）
        │  首次打开自动写入 IndexedDB
        ▼
浏览器 IndexedDB: datasets / records / profiles / entries / settings
        ▲
        │  「我的 → 导入考试总表」（真实课表，仅本机）
```

- 仓库内置的 `seed.json` 由示例脚本生成，不含任何真实记录。
- 如确需把真实总表做成内置数据：`node scripts/build-seed.mjs <总表路径>`，仅限本机或私有部署。
- 运行时解析用户新导入的 Excel 由浏览器内 `parser.js` 完成，文件不会上传。
- `seed.json` 只含结构化考试记录，不含原始表格、公式或教师信息。

## 目录结构

```
qingjianxu/
├── public/                  # 唯一发布根（含示例 seed.json）
├── data/source/             # 自有总表放这里（gitignore，不发布）
├── scripts/                 # 示例数据、可选真实种子、发布校验、本地服务器（不发布）
├── skills/                  # 命令行检索技能（不发布）
├── docs/screenshots/        # README 截图（不发布）
├── .github/workflows/       # Pages 部署（不发布）
├── .tools/                  # 图标与签名生成工具（不发布，gitignore）
└── README.md
```

## 隐私

| 场景 | 机制 |
| --- | --- |
| 仓库内置数据 | 全部为虚构示例，无真实院校、班级、课程或个人信息 |
| 导入真实总表 | 仅在浏览器内解析并写入 IndexedDB，不上传服务器 |
| 多份总表互不干扰 | `records` 以 `datasetId#n` 为主键，删除走主键区间 |
| 多人共用一台设备 | `profiles` 各自保存班级、关注、壁纸；总表可标记私有 |
| 跨设备 | 不做云端同步，用「导出数据包 / 导入数据包」传递 |
| 无痕模式 | IndexedDB 不可用时降级为内存存储，功能可演示但刷新即失 |

## 浏览器支持

Chrome / Edge / Safari / Firefox 近两年版本；移动端 Safari 与 Android Chrome 均已适配安全区、底部面板与文件导入。

## 许可

MIT © 2026 叶盗

考试时间与地点请以教务系统及学院通知为准。本工具仅作查询辅助，不对数据准确性作担保。
