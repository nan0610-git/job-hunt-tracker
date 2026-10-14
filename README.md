# job-hunt-tracker 🎯

> 求职投递管家 —— WorkBuddy 技能 | A Job Application Tracker & Offer Decision Skill for WorkBuddy

一站式管理求职全流程：**投递进度跟踪 → HR 跟进话术 → 多 Offer 加权对比决策 → 岗位监控与半自动辅助投递**。所有数据存在你自己的 Excel 表格里，不经过任何第三方。

## ✨ 功能特性

| 模块 | 能力 |
|------|------|
| 📋 投递跟踪 | Excel 双工作表（投递记录 / Offer对比），`add / update / list / stats` 四个命令管理投递生命周期 |
| 💬 跟进话术 | 8 个场景模板：催进度、面试感谢信、催结果、请求考虑时间、谈薪开场、婉拒 Offer、接受确认、改期 |
| ⚖️ Offer 对比 | 薪资包 40% / 发展前景 30% / 工作生活平衡 30% 加权打分，年薪自动归一化，分差决策规则（<0.3 打平 / >0.8 明确倾向） |
| 🔍 岗位监控 | Boss直聘 / 智联招聘自动筛选新岗位，六重过滤去重（关键词/薪资/黑名单/排除词/已见过/已投递） |
| 🤝 辅助投递 | 半自动设计：自动打开岗位页 + 预填打招呼语，**发送由用户点击**，投完自动记账 |

## 📦 安装

### 方式一：安装到 WorkBuddy（推荐）

将本仓库克隆为技能目录：

```bash
git clone https://github.com/zhangnan610/job-hunt-tracker.git ~/.workbuddy/skills/job-hunt-tracker
```

Windows 用户（Git Bash）：

```bash
git clone https://github.com/zhangnan610/job-hunt-tracker.git "C:/Users/<你的用户名>/.workbuddy/skills/job-hunt-tracker"
```

重启 WorkBuddy 后，对话中说「帮我记一笔投递」「看看BOSS上有没有新岗位」即可触发。

### 方式二：手动放置

下载仓库 zip，解压后把 `job-hunt-tracker` 文件夹放入：

- 用户级（所有项目可用）：`~/.workbuddy/skills/`
- 项目级（仅当前项目）：`<工作区>/.workbuddy/skills/`

### 依赖

脚本依赖 Python 3 + openpyxl：

```bash
pip install openpyxl
```

## 🚀 快速上手

对 WorkBuddy 说：

- 「帮我建个求职跟踪表」—— 初始化 Excel 跟踪表
- 「记一笔投递：XX公司 运维工程师，BOSS上投的，15-20K」—— 记录投递
- 「XX科技进二面了」—— 更新进度
- 「看下我的投递统计」—— 各状态数量 + 待跟进提醒
- 「XX公司一周没回我了」—— 生成催进度话术
- 「拿到两个offer帮我对比下」—— 引导录入 + 加权打分
- 「看看BOSS上有没有新的运维岗位」—— 自动筛选推送清单

## 🔒 隐私与合规

- **数据本地化**：投递记录、Offer 数据全部保存在本地 Excel，不采集、不上传。
- **平台协议**：Boss直聘 / 智联招聘的用户协议禁止自动化操作。本技能采用半自动设计——自动部分仅为「打开页面 + 读取内容」，点击发送永远由用户完成。请遵守平台规则，账号风险自担。
- **风控红线**：单日浏览 ≤ 50 岗位、遇到验证码立即转人工、不绕过任何平台验证机制。

## 📄 目录结构

```
job-hunt-tracker/
├── SKILL.md                        # 技能主指令（任务路由）
├── scripts/
│   ├── tracker_manager.py          # 跟踪表管理：init/add/update/list/stats/add-offer/compare
│   └── job_monitor.py              # 岗位监控：config/filter/seen/show/clear
├── references/
│   ├── followup-templates.md       # 8 场景 HR 跟进话术模板
│   ├── offer-evaluation.md         # Offer 评估方法论（一票否决 + 三维度打分）
│   └── job-monitor-guide.md        # 平台抓取与风控红线指南
└── README.md
```

---

用 WorkBuddy 的 skill-creator 制作 ✦ 由 [@zhangnan610](https://github.com/zhangnan610) 发布
