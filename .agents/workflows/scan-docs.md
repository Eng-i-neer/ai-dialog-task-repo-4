---
description: 扫描项目中的新文件并自动更新各目录的 README 文档
---

# 代码文档扫描 Workflow

执行以下步骤：

1. 读取 `docs/代码文档管理总控.md`，理解其中嵌入的 PROMPT 指引
2. 扫描项目根目录及所有子目录，获取所有文件列表（排除 `__pycache__`、`.pyc`、`.db`、`node_modules`、`.git` 等）
3. 读取各目录的 `README.md`，提取已记录的文件名列表
4. 对比差异，找出新增但未记录的文件、已删除但仍在记录中的文件
5. 对新文件使用 `view_file_outline` 和 `view_file` 分析功能
6. 将新文件的说明追加到对应目录的 `README.md`
7. 更新 `docs/代码文档管理总控.md` 底部的文件状态追踪表
8. 向用户汇报扫描结果
