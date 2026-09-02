# VBA 代码 Git 版本控制

本工程使用两套 Git 做隔离：

- **主仓库（本目录）**：只管理技能与文档（`.trae/`、`AGENTS.md`、`README.md`、`reset_vba_src.bat` 等），推送到 GitHub。
- **vba_src 子仓库**：独立本地 Git 仓库（无远程），存放每次修改 VBA 时导出的临时代码，仅供 AI 和你看差异，用完即弃。

## 项目简介

- 将 Excel 中的 VBA 代码临时导出为文本，便于查看每次修改的具体差异
- AI 在 Excel 中直接修改代码，再导出对比基线，可清晰回看变更内容
- 对比用的临时提交只留在本地 vba_src 子仓库，不会进入主仓库、不会推送到远程
- 任务结束后可一键清空重置 vba_src

## 目录结构

```
e:/python_space/vba_codes/
├── svn跨分支合表工具.xlsm          # Excel 文件（不入 Git，代码的真实宿主）
├── vba_src/                        # 独立本地 Git 子仓库（无远程，用完即弃）
│   └── <工作簿名>/                 # 导出的临时 VBA 代码
│       ├── Module1.bas
│       ├── Module2.bas
│       └── ...
├── reset_vba_src.bat               # 一键清空重置 vba_src 子仓库
├── .trae/skills/excel-vba-editor/  # AI 编辑 VBA 的技能与脚本（主仓库 Git 跟踪）
├── AGENTS.md                       # Agent 规则
└── README.md                       # 本文件
```

## 如何使用

### 1. 让 AI 帮你修改代码（推荐）

直接告诉 AI 你的需求，并显式调用技能，例如：

```
调用技能excel-vba-editor，修改 svn跨分支合表工具 中的VBA代码
```

然后，AI 会：
1. 查看当前 VBA 代码，确定要改的模块
2. 导出基线代码到 vba_src 并提交（本地快照）
3. 直接在 Excel 中修改 VBA 代码
4. 再次导出，用 vba_src 内的 `git diff` 向你展示变更

你只需在 Excel 中保存文件即可。

### 2. 查看差异 / 历史（在 vba_src 内）

打开终端并执行：

```
cd e:/python_space/vba_codes/vba_src
git log --oneline     # 查看本次任务的提交历史
git diff HEAD~1       # 查看最近一次修改差异
```

### 3. 任务结束后清空 vba_src

vba_src 只是临时对比仓库，没有长期保留价值。双击根目录的 `reset_vba_src.bat`（或终端执行 `reset_vba_src.bat auto` 免确认），它会删除整个 vba_src 并重建为全新本地仓库。清空不影响主仓库与 Excel 文件。

## 注意事项

1. **Excel 文件不入 Git**：`.xlsm` 的权威代码只保存在 Excel 中，请经常在 Excel 里 Ctrl+S 保存。
2. **vba_src 不入主仓库**：它是独立本地子仓库（无远程），主仓库 `.gitignore` 已忽略该目录，其中的提交不会被推送。
3. **修改前确保 Excel 已打开**：AI 需要连接正在运行的 Excel 实例。
4. **请勿在主仓库 `git add vba_src`**：该目录已被忽略，强行添加会破坏子仓库的独立性。
