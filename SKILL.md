---
name: job-hunt-tracker
description: 求职投递跟踪、岗位监控与Offer决策管理技能。当用户需要记录或查询求职投递进度、统计投递状态、生成HR跟进话术（催进度/面试感谢/谈薪/婉拒Offer）、对多个Offer进行对比决策、自动筛选监控新岗位（Boss直聘/智联招聘）、或请求半自动辅助投递时使用。触发词：投递记录、求职跟踪、投递进度、跟进话术、催HR、面试感谢信、offer对比、拿了好几个offer怎么选、岗位监控、帮我找新岗位、自动投递、job hunt、application tracker。
---

# Job Hunt Tracker（求职投递管家）

## Overview

本技能提供求职全流程管理能力：投递进度跟踪（Excel 数据表持久化）、HR 跟进话术生成、多 Offer 多维度对比决策。所有数据读写通过 `scripts/tracker_manager.py` 完成，保证表格格式一致。

## 快速开始

1. 确定数据文件路径：默认使用当前工作目录下的 `求职跟踪表.xlsx`；用户指定了路径则以用户为准。
2. 文件不存在时先执行初始化（任务一）。
3. 按用户需求执行对应任务：记录投递 / 查询统计 / 生成话术 / Offer 对比。

脚本通用调用方式（依赖 openpyxl，缺失时提示 `pip install openpyxl`）：

```
python "<本技能目录>/scripts/tracker_manager.py" <子命令> [参数]
```

## 任务一：初始化跟踪表

```
python tracker_manager.py init [文件路径] [--force]
```

- 生成两个工作表：「投递记录」与「Offer对比」，含表头样式、状态下拉选项、冻结首行。
- 文件已存在时默认拒绝覆盖；仅当用户明确要求重建时才加 `--force`（重建前提醒用户旧数据将丢失）。

## 任务二：记录与更新投递

```
python tracker_manager.py add <文件> --company "XX公司" --role "运维工程师" [--location 城市] [--channel BOSS直聘] [--date 2026-10-14] [--salary "15-20K"] [--status 投递中] [--followup 2026-10-21] [--note 备注]
python tracker_manager.py update <文件> --company "XX公司" [--role 岗位] [--status 二面] [--note "面试官强调K8s经验"]
python tracker_manager.py list <文件> [--status 二面]
python tracker_manager.py stats <文件>
```

行为要点：
- 用户未提供投递日期时默认今天；未指定跟进日期时默认投递后 7 天，并在输出中告知。
- 状态枚举固定为：投递中 / 笔试 / 一面 / 二面 / HR面 / Offer / 已拒绝 / 已婉拒 / 无回音。
- `update` 按公司名定位，同名公司多岗位时要求用户提供 `--role` 区分。
- `stats` 输出总投递数、各状态计数、本周新增、待跟进清单；主动播报风险信号（如「无回音」占比超 40% 时建议优化简历或拓宽渠道）。

## 任务三：生成跟进话术

先读取 `references/followup-templates.md`，按场景匹配模板生成。场景与模板对应关系：

| 场景 | 模板 |
|------|------|
| 投递后约一周无回音 | 催进度 |
| 面试结束当天 | 面试感谢信 |
| 面试后超一周无结果 | 催结果 |
| 收到口头 Offer 想争取考虑时间 | 请求考虑时间 |
| 想提升薪资包 | 谈薪开场 |
| 拒绝其他 Offer | 婉拒话术 |
| 确认接受 Offer | 接受确认 |

生成要求：用跟踪表中的公司/岗位/进度信息填充占位符；语气礼貌专业、不卑不亢；一条消息只说一件事。输出后提醒用户可要求调整语气（更正式 / 更亲切 / 更简短）。

## 任务四：Offer 对比决策

1. 读取 `references/offer-evaluation.md` 掌握评估方法论与打分标准。
2. 录入 Offer 数据（缺信息时逐项向用户询问，不要编造数字）：

```
python tracker_manager.py add-offer <文件> --company "XX公司" --role "运维工程师" [--month-salary 18000] [--annual 30] [--stock 10] [--bonus-months 3] [--social-insurance "全额12%"] [--scores "4,4,5,4,3"] [--note 备注]
```

`--scores` 顺序固定为：技术匹配, 成长空间, 稳定性, 工作轻松度, 通勤便利（每项 1-5）。

3. 运行对比：

```
python tracker_manager.py compare <文件> [--weights 0.4,0.3,0.3]
```

4. 结合脚本输出的加权排序表给出结论：先确认无一票否决项，再按分差规则解读（<0.3 打平、0.3-0.8 倾向高分者、>0.8 明确倾向）；列出每个 Offer 的最强项与最弱项，而不是只报分数。
5. 权重默认：薪资包 40%、发展前景 30%、工作生活平衡 30%；用户表达明显偏好时按其偏好调整 `--weights` 并说明。

## 任务五：岗位监控与辅助投递（半自动）

对 Boss直聘 / 智联招聘的符合条件岗位做「自动筛选 + 辅助投递」。**先读取 `references/job-monitor-guide.md`** 了解抓取流程、DOM 要点与风控红线，再执行。

1. 读取配置（首次使用先询问用户的筛选条件，然后初始化）：

```
python "<本技能目录>/scripts/job_monitor.py" config [--init] [--set key=value ...]
```

配置项：keywords（岗位关键词）、city、salary_min_k（薪资下限K）、exclude_companies（黑名单公司）、exclude_keywords（排除词如"外包"）、daily_limit（每日上限）。

2. 用浏览器能力（playwright-cli / workbuddy-browser-use）按配置打开平台搜索页，解析岗位卡片保存为 `raw_jobs.json`（字段格式见 guide）。
3. 过滤去重，生成待投清单：

```
python "<本技能目录>/scripts/job_monitor.py" filter raw_jobs.json --tracker <跟踪表.xlsx> --data-dir <数据目录>
```

自动完成：关键词/薪资/黑名单过滤、与历史已见岗位去重（seen_jobs.json）、与跟踪表已投递记录去重。输出按薪资降序的清单。

4. 把清单列给用户确认。**未经用户确认绝不进入投递步骤。**
5. 用户确认后逐个辅助投递（流程见 guide）：打开岗位页 → 预填打招呼语 → **用户自己点击发送** → `tracker_manager.py add` 记账 → `job_monitor.py seen --url <链接>` 标记已处理。
6. 待投清单可用 `show` 查看、`seen --all` / `clear` 管理。
7. 用户需要每日推送时，建议创建每天上午的 automation：只执行抓取+过滤+呈现清单（guide 第五节）。

## 注意事项

- 所有对 Excel 的修改通过脚本完成，避免直接编辑破坏格式。
- 薪资保留用户原始表述，换算只在 Offer 对比内部进行。
- 投递记录属敏感个人数据，不要上传到任何外部服务。
- 平台协议禁止自动化操作：始终遵守 guide 中的风控红线（频率限制、验证码交人工、发送由用户点击），用户坚持全自动时告知封号风险且不绕过任何验证机制。
