## 项目简介

本仓库（主仓库）是基于 VBA 的 Excel 工具的配套工程，用 Git 管理**技能与文档**，并支撑"修改 Excel VBA 后查看差异"的流程：

- **业务功能**：`svn跨分支合表工具.xlsm` 用 VBA 比较两个 Excel 工作簿的差异。
- **主仓库内容**：`.trae/skills/excel-vba-editor/` 技能、`AGENTS.md`、`README.md`、`reset_vba_src.bat` 等，推送到 GitHub。
- **VBA 代码载体**：真正的代码只存在于打开的 Excel（`.xlsm`）中，仓库内没有可独立运行的副本。

## Agent 核心认知与规则

1. **代码的真正宿主是打开的 Excel 文件。**
   - 用户要你修改的 VBA 代码位于 `svn跨分支合表工具.xlsm`（二进制，不入任何 Git 仓库）。
   - `vba_src/` 下导出的 `.bas` / `.cls` / `.frm` 只是**临时镜像**，供 diff 对比用，不是权威源码。
2. **修改 VBA 必须调用 `excel-vba-editor` 技能**，直接编辑已打开的 Excel 实例；不要直接改动 `vba_src` 里的文本后让用户手动同步。
3. **修改前置条件**：目标工作簿必须已在 Excel 中打开，否则无法连接实例写回代码。
4. **vba_src 是独立本地子仓库（无远程），用完即弃**：
   - 它不属于主仓库：主仓库的 `.gitignore` 已忽略 `vba_src/`，其内任何提交都不会进入主仓库、不会推送到 GitHub。
   - 临时对比 commit 一律在 vba_src **内部**执行：先 `cd vba_src`，再 `git add .` / `git commit` / `git diff`，切勿在 vba_src 内 push 或给主仓库 `git add vba_src`。
   - 一次任务结束后，该子仓库的历史没有保留价值：用户双击根目录 `reset_vba_src.bat`（或让 agent 代跑 `reset_vba_src.bat auto` 免确认）即可整目录删除并重置为全新仓库，**不要把它当作长期代码历史库**。
5. **改动完成后**：提醒用户在 Excel 中保存文件，用 vba_src 内的 `git diff` 展示变更摘要；收尾时询问用户是否需要重置 vba_src。

## 目录结构（主仓库视角）

| 路径 | 说明 |
| --- | --- |
| `svn跨分支合表工具.xlsm` | Excel 主文件（不入 Git，VBA 代码的实际宿主） |
| `vba_src/` | 独立本地 Git 子仓库（无远程；主仓库已 gitignore；临时导出，用完即弃） |
| `reset_vba_src.bat` | 一键清空并重置 vba_src 子仓库（支持 `auto` 参数免确认） |
| `.trae/skills/excel-vba-editor/` | excel-vba-editor 技能（含脚本），主仓库 Git 跟踪 |
| `README.md` | 面向人类用户的使用说明 |
| `.temp/` | 临时文件、临时脚本（Git 忽略） |

**代码定位规则**：导出后，工作簿对应 `vba_src/<工作簿名>/`，模块对应子目录内单文件——`vba_src/<工作簿名>/<模块名>.bas`（类模块 `.cls`、窗体 `.frm` 同理）。

## Agent 修改 VBA 代码的标准流程

1. 用 `excel-vba-editor` 技能的查看脚本（`list_modules.py` / `read_module.py`）理解现有代码与结构。
2. 确认 Excel 已打开该工作簿；在 vba_src 内导出基线代码并提交：`cd vba_src; git add .; git commit -m "导出原始代码 - <模块名>"`。
3. 用技能脚本直接编辑 Excel 中的代码，再导出到 vba_src。
4. 在 vba_src 内执行 `git diff` 展示变更摘要，等待用户在 Excel 中确认保存。
5. 收尾询问用户是否运行 `reset_vba_src.bat` 清理临时仓库。

## 开发环境

- Excel 2019 以上 / VBA 7.0 / Windows 11 以上
- Python 3.11 以上（供 agent 调用辅助脚本）
