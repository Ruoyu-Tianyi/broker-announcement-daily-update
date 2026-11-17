# 文件级工作簿引擎

当运行环境提供 Python 3 和 openpyxl 3.1.5 或更高版本时，优先使用 `scripts/workbook_engine.py`。该脚本只读写文件，不控制用户桌面应用。

专属工作区的初始化、日期批次、输入归档和活动总表晋升使用 `scripts/workspace_manager.py`。先按 [专属工作区与总表生命周期](workspace-layout.md) 建立/读取批次，再调用工作簿引擎；工作簿引擎的输入输出不得绕过当前 `run.json`。

## 使用顺序

1. 先运行 `selftest`，确认整行插入、历史区校验、来源对账和富文本往返均通过。
2. 对真实文件运行 `preflight`。只有输出状态为 `ready` 且 `blockers` 为空时才可继续。
3. 运行 `stage1` 生成新的输出文件和 JSON 报告。输出路径必须不存在，脚本拒绝覆盖。
4. 结合文件级渲染检查输出中新区域及上下边界。
5. 阶段二由 Agent 逐份阅读 PDF 并生成带来源路径的决议 JSON，再运行 `stage2`。脚本会拒绝缺项、重复单元格、标题不一致或来源 PDF 不存在的输入，并只允许目标单元格及对应新增行行高发生变化。

```text
python scripts/workbook_engine.py selftest
python scripts/workbook_engine.py preflight --master 总表.xlsx --source 昨晚.xlsx --source 今早.xlsx --report 预检.json
python scripts/workbook_engine.py stage1 --master 总表.xlsx --source 昨晚.xlsx --source 今早.xlsx --output 阶段一.xlsx --report 阶段一.json
python scripts/workbook_engine.py stage2 --input 阶段一.xlsx --stage1-report 阶段一.json --resolutions 决议.json --output 最终.xlsx --report 阶段二.json
```

## 安全设计

- 自动识别所有日更区，按区块内最大公告日期选择最新的完整区块，不以行号或文件修改时间猜测。
- 新表头复制时同时重建原表题的横向合并；新增日期单元格使用本 Skill 范例规定的 `yyyy/m/d` 显示格式。
- 检查公式、定义名称、外部链接、结构化表、条件格式、数据验证、图表、图片、宏、透视表、切片器等 openpyxl 无法安全迁移的特性；发现即阻止写入。
- 在插入前记录历史区和其他工作表的值、公式、样式、批注、超链接、合并关系、行列尺寸及页面设置；保存并重载后逐项比较。
- openpyxl 会丢弃部分打印机设置部件。脚本会从原总表恢复对应 OOXML 部件、关系和内容类型，并检查是否仍有非允许部件丢失。
- 允许 shared strings 转换为 inline strings，因为其显示值和单元格语义不变；其他意外 OOXML 部件丢失会使任务失败。
- 阶段一报告包含来源行数、插入范围、合并范围、董事会清单及全部验证结果，供阶段二定位与审计。
- 阶段二要求决议 JSON 中每项同时包含 `broker`、`title`、`sheet`、`cell`、`pdf` 和 `items`；保存重载后验证标题富文本加粗、正文非粗体、自动换行、编号完整，以及目标之外的全工作簿内容和格式不变。

如果 `preflight` 返回阻止项，不得修改脚本以绕过检查；应报告具体特性，并改用另一种同样是文件级、能够无损处理该特性的工具。
