# 专属工作区与总表生命周期

本规范确保每日输入、输出和总表版本可追溯，同时始终只有一份明确的活动总表。

## 1. 首次调用必须初始化

若当前任务没有可识别的 `workspace.json`，先完整阅读并执行 [首次使用介绍与初始化交互](first-use.md)，暂停日更处理并请求用户确认两项：

1. 专属工作区根目录；
2. 初始【日更】券商公告总表。

不要把任意当前目录静默当成专属工作区，也不要移动或删除用户提供的初始总表。确认后运行 `scripts/workspace_manager.py init`，将初始总表复制为活动总表并建立目录。初始化只执行一次；后续调用直接读取 `workspace.json`。

```text
工作区/
├─ workspace.json
├─ master/
│  └─ 【日更】券商公告.xlsx
├─ input/
│  └─ YYYY-MM-DD/
│     └─ run-001/
│        ├─ 01_昨晚附件.xlsx
│        ├─ 02_今早附件.xlsx
│        ├─ pdf/
│        └─ run.json
├─ output/
│  └─ YYYY-MM-DD/
│     └─ run-001/
│        ├─ 阶段一.xlsx
│        ├─ 最终.xlsx
│        └─ 验证报告.json
└─ archive/
   └─ master/
      └─ YYYY-MM-DD/
         └─ run-001/
            └─ 【日更】券商公告_before_YYYY-MM-DD_run-001.xlsx
```

目录名 `input`、`output`、`master`、`archive` 和配置文件名固定；工作区根目录由用户决定。

## 2. 日期和批次

- 归档日期使用本批公告的业务日期，通常是两份来源 Excel 中较新的有效公告日期；不要使用 PDF 实际上传日替代。
- 同一批 PDF 即使在次日补传，也归入原 Excel 批次的 `input/YYYY-MM-DD/run-NNN/pdf/`。
- 同一业务日期的首次处理是 `run-001`。补充或重新开始的新批次依次使用 `run-002`、`run-003`，不得覆盖既有批次。
- `start-run` 会记录两份来源 Excel 的 SHA-256；完全相同的文件对已经归档时停止，防止重复插入总表。

## 3. 输入归档

收到当次对话明确上传或点名的两份 Excel 后，先运行 `start-run`，把它们复制进本批 `input` 目录，再以归档副本作为阶段一输入。后续 PDF 运行 `attach-pdf` 复制到同一批次。

归档不改变来源资格：只有本次对话明确上传或点名的文件才能启动或补充当前批次；历史 `input` 中的文件只能用于审计和用户明确要求的重跑，不能自动充当今天的附件。

`run.json` 至少记录：业务日期、批次号、来源原路径、归档路径、文件哈希、活动总表处理前哈希、PDF 清单、输出目录和状态。

## 4. 输出归档

- 阶段一、阶段二和验证报告都写入当前批次的 `output/YYYY-MM-DD/run-NNN/`，文件名稳定且不覆盖既有文件。
- 有待补董事会决议时，阶段一文件只是中间结果，不能替换活动总表。
- 没有待补项时，通过全部验收的阶段一文件可视为最终文件。
- 有待补项时，只有全部 PDF 处理完成、`未解决项目数=0` 且最终验收通过的阶段二文件才是最终文件。
- 校验失败或资料不完整时保留当前输出和报告，批次不晋升，活动总表保持不变。

## 5. 活动总表更新

每个完整批次都从 `master/【日更】券商公告.xlsx` 开始，并在最终验收通过后自动晋升最终文件为新的活动总表，无需每天再次询问是否采用。晋升顺序不可改变：

1. 确认活动总表哈希仍等于 `run.json` 中的处理前哈希；不一致说明有并发或人工修改，停止晋升。
2. 把当前活动总表完整复制到 `archive/master/YYYY-MM-DD/run-NNN/`。
3. 把最终文件复制到活动总表同目录的临时文件，核对哈希。
4. 使用原子替换更新 `master/【日更】券商公告.xlsx`。
5. 重新核对活动总表哈希，并更新 `workspace.json` 和 `run.json`。

禁止先覆盖总表再备份。禁止删除历史输入、输出或总表备份。用户要求清理时，把删除作为独立任务，明确列出目标后再执行。

## 6. 命令

```text
python scripts/workspace_manager.py init --root 工作区 --master 初始总表.xlsx
python scripts/workspace_manager.py start-run --root 工作区 --date YYYY-MM-DD --source 昨晚.xlsx --source 今早.xlsx
python scripts/workspace_manager.py attach-pdf --root 工作区 --date YYYY-MM-DD --run run-001 --pdf 公告1.pdf --pdf 公告2.pdf
python scripts/workspace_manager.py promote --root 工作区 --date YYYY-MM-DD --run run-001 --final 最终.xlsx --report 验证报告.json
```

运行 `init`、`start-run` 或 `promote` 前先确认参数中的绝对路径与预期工作区一致。脚本拒绝覆盖现有工作区配置、活动总表、批次目录、归档 PDF、最终输出和旧总表备份。
