---
name: "excel-vba-editor"
description: "使用 xlwings 直接编辑已打开的 Excel .xlsm 文件中的 VBA 代码。支持导出代码用于Git差异对比，直接在Excel中编辑，再导出对比修改内容的完整工作流。"
---

# Excel VBA 编辑器

使用 xlwings 直接编辑已打开的 Excel 文件 (.xlsm) 中的 VBA 代码。

本技能提供**直接编辑**工作流，导出的文件仅用于Git差异对比：
- **分析**：确定需要修改的模块（标准模块、类模块等）
- **导出**：将代码导出到 `vba_src/` 目录（仅用于查看差异）
- **版本控制**：Git commit 保存原始代码快照
- **直接编辑**：AI直接在Excel工作簿中修改VBA代码
- **再导出**：编辑后再次导出，查看修改差异

## 前提条件

1. Excel 正在运行且目标 .xlsm 文件已打开
2. 已安装 xlwings：`pip install xlwings`
3. Excel 中启用"信任对 VBA 项目对象模型的访问"（信任中心 > 宏设置）

## 推荐工作流（直接编辑模式）

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│  1. 导出代码     │───▶│  2. Git commit  │───▶│  3. 直接编辑     │
│  (用于对比基线)  │    │  (保存原始快照)  │    │  (在Excel中修改) │
└─────────────────┘    └─────────────────┘    └─────────────────┘
                                                        │
┌─────────────────┐    ┌─────────────────┐             │
│  5. 查看差异     │◀───│  4. 再导出代码   │◀────────────┘
│  (Git diff对比)  │    │  (查看修改内容)  │
└─────────────────┘    └─────────────────┘
```

### 快速开始

```bash
# 1. 分析任务，确定需要修改的模块
python .trae\skills\excel-vba-editor\scripts\list_modules.py "workbook.xlsm"

# 2. 导出需要修改的模块代码到 vba_src/（作为对比基线）
python .trae\skills\excel-vba-editor\scripts\export_vba.py "workbook.xlsm" "ModuleName"

# 3. Git commit 保存原始代码快照（必做！）
# PowerShell 或 CMD 通用命令：
git add "vba_src/workbook/"
git commit -m "[excel-vba-editor]导出原始代码 - ModuleName"

# 4. 【关键】直接在Excel中编辑VBA代码
#    AI使用 modify_module.py / modify_method.py 等脚本直接修改工作簿

# 5. 编辑完成后，再次导出查看修改内容
python .trae\skills\excel-vba-editor\scripts\export_vba.py "workbook.xlsm" "ModuleName"

# 6. 使用Git diff查看修改差异
git diff vba_src/
```

### 工作流说明

1. **分析修改范围**：明确本次任务需要修改哪些模块
2. **导出代码（基线）**：导出原始代码到 `vba_src/`，仅用于后续对比差异
3. **Git 提交**：**必须立即执行**！提交原始代码快照，建立对比基线
4. **直接编辑**：AI使用脚本直接在Excel工作簿中修改VBA代码，**不编辑导出的文件**
5. **再导出（对比）**：编辑完成后再次导出，用于查看实际修改内容
6. **查看差异**：使用 `git diff` 对比两次导出的差异，确认修改内容

**核心原则**：导出的文件仅用于Git差异对比，实际编辑直接在Excel中进行。

**⚠️ Git Commit 强制要求**：
- 在导出代码后、编辑代码前，必须执行 Git commit 保存原始代码快照
- 这是后续对比修改差异的唯一基线，不可省略
- 如果跳过此步骤，将无法准确查看代码变更内容
- PowerShell/CMD 通用命令示例：
  ```bash
  git add "vba_src/工作簿名称/"
  git commit -m "[excel-vba-editor]导出原始代码 - 模块名"
  ```

## ⚠️ 重要提示

**导出文件仅供查看差异，不要直接编辑！**

正确的工作方式：
1. 导出 → Git commit → **直接在Excel中编辑** → 再导出 → Git diff查看差异
2. 不是：导出 → 编辑文件 → 导入 → 写回Excel

如果用户要求修改代码，**直接在Excel工作簿中进行修改**，而不是编辑导出的文件。

## 直接编辑脚本

| 脚本 | 功能 | 用法示例 |
|------|------|----------|
| `edit_module.py` | **推荐：直接编辑模块** | `python scripts\edit_module.py "wb.xlsm" "Mod1" "FILE:code.bas"` |
| `modify_module.py` | 替换整个模块代码 | `python scripts\modify_module.py "wb.xlsm" "Mod1" "FILE:code.bas"` |
| `modify_method.py` | 修改特定方法/过程 | `python scripts\modify_method.py "wb.xlsm" "Mod1" "MySub" "FILE:code.vba"` |
| `modify_code_block.py` | 代码块替换（搜索替换） | `python scripts\modify_code_block.py "wb.xlsm" "Mod1" "old_code" "new_code"` |
| `add_module.py` | 添加新模块 | `python scripts\add_module.py "wb.xlsm" "NewMod" "FILE:code.bas"` |
| `delete_module.py` | 删除模块 | `python scripts\delete_module.py "wb.xlsm" "Module1"` |

**推荐使用 `edit_module.py`**：它是 `modify_module.py` 的封装，提供更友好的错误提示和使用说明。

**所有脚本支持 `FILE:路径` 前缀从文件读取代码。**

## 辅助脚本（用于查看和导出）

| 脚本 | 功能 | 用法示例 |
|------|------|----------|
| `list_workbooks.py` | 列出打开的工作簿 | `python scripts\list_workbooks.py` |
| `list_module_names.py` | 列出模块名称 | `python scripts\list_module_names.py "wb.xlsm"` |
| `list_modules.py` | 读取所有模块代码 | `python scripts\list_modules.py "wb.xlsm"` |
| `read_module.py` | 读取特定模块 | `python scripts\read_module.py "wb.xlsm" "Module1"` |
| `export_vba.py` | 导出到 vba_src/（用于对比） | `python scripts\export_vba.py "wb.xlsm" "Module1"` |

## 组件类型

| 类型 | 说明 | 扩展名 |
|------|------|--------|
| Standard Module | 标准模块 | .bas |
| Class Module | 类模块 | .cls |
| Form | 用户窗体 | .frm |
| Workbook/Worksheet Code | 工作簿/表代码 | .vba |

## 错误处理

| 错误 | 解决方案 |
|------|----------|
| "No active Excel application" | 启动 Excel 并打开 .xlsm 文件 |
| "Workbook not found" | 使用 `list_workbooks.py` 查看可用工作簿 |
| "Module not found" | 使用 `list_module_names.py` 查看可用模块 |
| "Error accessing VBA project" | 在 Excel 中启用 VBA 项目访问信任 |

## 相关文档

- **USAGE.md** - 详细使用指南
- **scripts/** - 项目根目录辅助脚本
