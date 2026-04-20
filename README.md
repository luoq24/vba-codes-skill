# VBA 代码 Git 版本控制

本项目使用 Git 对 Excel VBA 代码进行版本控制，让你能够清晰查看代码变更历史，随时回退到之前的版本。

## 项目简介

- 将 Excel 中的 VBA 代码导出为文本文件，便于 Git 版本控制
- 支持查看每次修改的具体内容
- 可随时回退到任意历史版本

## 目录结构

```
e:/python_space/vba_codes/
├── svn跨分支合表工具.xlsm          # Excel 文件（不提交到Git）
├── vba_src/                        # 导出的 VBA 代码目录（Git 跟踪）
│   └── svn跨分支合表工具/          # 工作簿专属目录
│       ├── Module1.bas
│       ├── Module2.bas
│       └── ...
└── README.md                       # 本文件
```

## 如何使用

### 1. 查看 VBA 代码变更

所有 VBA 代码都会导出到 `vba_src/` 目录下：

- 每个工作簿对应一个子目录
- 每个模块对应一个 `.bas`、`.cls` 或 `.frm` 文件

**在 Trae CN 中查看代码变更：**
1. 打开 `vba_src/工作簿名称/模块名.bas` 文件
2. 在左侧 Git 面板查看修改内容
3. 或使用 `git diff` 命令查看差异

### 2. 让 AI 帮你修改代码

直接告诉 AI 你的需求。最好显示的调用skill，例如：
```
调用技能excel-vba-editor，修改 @ 中的VBA代码
```
然后，AI 会：
1. 查看当前的 VBA 代码
2. 进行修改
3. 将修改后的代码写回 Excel

你只需在 Excel 中保存文件即可。

## 注意事项

1. **Excel 文件不纳入版本控制**：Git 只跟踪 `vba_src/` 目录下的文本格式 VBA 代码
2. **修改前确保 Excel 已打开**：AI 需要连接到正在运行的 Excel 实例
3. **定期查看 Git 变更**：建议在 AI 修改后查看 `git diff`，确认修改内容
