---
name: qingjianxu-exam-finder
description: 从青简叙内置种子或导出的数据包中，按年级、专业、班级、课程、日期检索考试时间与考场，输出对齐表格或 JSON。当用户询问“我们班什么时候考某门课”“下周有哪些考试”“我的考场在哪”时使用。
license: MIT
metadata:
  version: "1.0.0"
  language: zh-CN
---

# qingjianxu-exam-finder · 考试信息查找

面向已有规范化考试数据的确定性检索，不做猜测。

## 何时使用

- 用户给出班级 / 年级 / 专业 / 课程名 / 日期，想查考试时间、考场、场次。
- 需要判断「下一场考试还有几天」「这周有几场」。
- 需要把检索结果交给下游继续处理。

## 数据来源

| 来源 | 路径 | 说明 |
| --- | --- | --- |
| 仓库内置种子 | `public/assets/data/seed.json` | 由 `npm run seed` 生成的虚构示例数据 |
| 用户导出数据包 | 浏览器「我的 → 导出数据包」得到的 `*.json` | 含完整记录，可离线分发 |

记录字段：

```json
{
  "id": "ds-demo-2026-2027-1#1",
  "course": "中国古代文学（三）",
  "code": "QJ-HY2401",
  "date": "2026-11-16",
  "start": "09:00",
  "end": "11:00",
  "grade": "24级",
  "major": "汉语言",
  "classNo": "01",
  "cls": "24级汉语言01班",
  "classKey": "24级|汉语言|01",
  "rooms": ["博雅楼 A201"],
  "teaching": "24级汉语言01班;24级汉语言02班",
  "headcount": 32,
  "iso": "2026-11-16T11:00:00"
}
```

## 用法

```bash
node skills/exam-finder/find.mjs --class "24级汉语言01班"
node skills/exam-finder/find.mjs --grade 26级 --major 计科 --upcoming
node skills/exam-finder/find.mjs --course 数据结构 --json
node skills/exam-finder/find.mjs --class "25级法学01班" --dataset 0
node skills/exam-finder/find.mjs --data "/path/to/青简叙数据包_xxx.json" --upcoming
```

参数：`--class` `--grade` `--major` `--course` `--date` `--dataset` `--upcoming` `--limit` `--json` `--data`。

## 输出约定

- 默认输出等宽对齐表格，中文按双宽计算；
- `--json` 输出 `{ count, records }`，便于程序消费；
- 无匹配时退出码为 1，并在 stderr 说明。

## 注意

- 自建条目（手动录入）只存在浏览器 IndexedDB，不写入 `seed.json`；命令行看不到，除非用户导出数据包时另行提供条目数据。
- 数据时间以总表为准；如与教务系统不一致，以教务系统为准。
