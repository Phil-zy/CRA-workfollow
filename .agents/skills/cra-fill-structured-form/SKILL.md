---
name: cra-fill-structured-form
description: Safely fill one CRA structured Word or Excel form from confirmed project facts and generate a review checklist. Use for CRA requests such as 填写伦理审查申请表、根据项目信息填写 Word 或 Excel 表单、生成待核对表单或填写核对清单 when the target is one .docx or .xlsx file and the facts come from an approved normalized local workbook. Require workspace authorization and human approval for every new or changed template mapping. Do not use for subject source data, direct identifiers, automatic submission or signing, batch filling, or raw-document ingestion.
---

# CRA 通用结构化表单填写

把本 Skill 作为用户入口。只把 `scripts/` 中的 Python 文件作为内部确定性执行能力；不要要求 CRA 直接运行脚本。

## 首版边界

- 每次只处理一个 `.docx` 或 `.xlsx` 目标表单。
- 只读取获准工作区内的规范化 `.xlsx` 项目事实。
- 不解析原始项目资料，不执行 OCR，不连接知识库或外部服务。
- Excel 只填写模板配置明确指定的普通单元格或合并区域左上角单元格；不覆盖公式。
- 不发送、上传、签署、归档或覆盖正式文件。

## 运行前必须读取

1. 读取 [safety-and-approval.md](references/safety-and-approval.md)，执行数据边界和审批规则。
2. 读取 [data-contracts.md](references/data-contracts.md)，校验事实文件、模板配置和输出契约。
3. 读取 [status-and-mapping-rules.md](references/status-and-mapping-rules.md)，应用六类状态和映射规则。
4. 在生成输出前读取 [output-validation.md](references/output-validation.md)，按目标格式执行结构与视觉验证。

## 执行流程

### 1. 收集路径但不读取内容

要求 CRA 明确提供或确认授权项目工作区、一个 `.docx` 或 `.xlsx` 目标文件、一个规范化 `.xlsx` 项目事实文件、输出目录和模板配置库。未指定输出目录时使用目标文件所在工作区的 `output/日期时间/`；未指定配置库时使用仓库根目录 `template-configs/cra-fill-structured-form/`。

解析绝对路径并确认全部输入位于授权工作区。不要扫描相邻目录或其他磁盘。

### 2. 展示授权摘要并等待确认

在读取文件内容前展示：目标文件、事实来源、输出目录、配置库、将使用的本地 Python 脚本，以及“联网：否；外部服务：否”。要求 CRA 明确确认目录已获授权、输入不含受试者源数据或直接识别信息，并允许 Agent 处理这些本地文件。未获得明确确认时停止。

### 3. 检查模板配置

调用 `scripts/inspect_template.py` 检查保护、宏、外部数据连接和可读性，并计算文件哈希和结构指纹。只允许以下分支：

- 存在状态为 `enabled` 且结构指纹、文件 SHA-256 均完全匹配的配置：继续读取事实。
- 没有匹配配置、只有草稿配置或模板结构已变化：调用 `scripts/draft_mapping.py` 生成新草稿，向 CRA 展示所有字段映射，然后停止。Excel 新模板还必须提供逐字段的工作表与单元格位置说明；核心程序不得写死具体 Excel 版式。

不要在生成映射草稿的同一审批步骤中填写表单。

### 4. 审批映射

只有 CRA 对草稿中的目标位置、事实键、字段类型、选项规则及人工保留区域作出明确批准后，才调用 `scripts/activate_mapping.py` 生成不可原地修改的已启用版本。不要修改已启用配置；任何变更都生成更高版本的新草稿，旧版本只能停用，不能删除。

### 5. 生成待核对输出

按目标格式调用 `scripts/fill_docx.py` 或 `scripts/fill_xlsx.py`，再次验证配置 Schema、审批状态、结构指纹和文件 SHA-256，只读输入并写入新的表单副本，对每个目标字段产生且只产生一种状态，把临时审计 JSON 写入任务临时目录。

随后调用 `scripts/build_checklist.py` 生成独立 `.xlsx` 核对清单。不得在正式表单中添加批注、问题标记或审核痕迹。

### 6. 验证并交付

Word 调用 `scripts/validate_outputs.py`；Excel 调用 `scripts/validate_xlsx_outputs.py`。验证输入及配置哈希、目标结构、实际填写、人工保留区域、状态集合、核对清单逐行内容及两个输出的审计关联，再按 [output-validation.md](references/output-validation.md) 完成视觉检查。

只有全部检查通过后，才能交付 `原文件名_待核对_YYYYMMDD-HHMMSS.原扩展名` 和 `原文件名_填写核对清单_YYYYMMDD-HHMMSS.xlsx`。明确说明输出仍需 CRA 审核，不是正式记录。

### 7. 记录 CRA 人工核对结果

CRA 对双输出作出明确核对结论后，调用 `scripts/finalize_review.py` 生成新的 `原文件名_CRA核对记录_YYYYMMDD-HHMMSS.xlsx` 和独立人工核对审计 JSON。不得修改或覆盖原待核对表单、原核对清单或执行审计。

随后调用 `scripts/validate_review_record.py`，重新验证输入、配置、表单输出、原核对清单、已核对清单的路径与 SHA-256，并确认每个字段的 `CRA 最终决定`、核对人、核对时间、核对 ID 和总体结论可追溯。只有验证通过后，才能把业务验收记为完成；签名、发送、提交和正式归档仍由 CRA 执行。

## 异常停止

发现以下任一情况时立即停止，不生成可能损坏或误填的结果：

- 疑似受试者源数据、姓名、身份证号、联系方式、住址或病历号。
- 输入不在授权工作区，文件权限不明，或目标为受保护、加密、含宏、含外部数据连接或非 `.docx`/`.xlsx` 文件。
- 模板配置未启用、结构指纹不匹配或映射审批记录缺失。
- 文件无法安全读取、复制或保存。
- 输出路径与任一输入路径相同。

单个字段缺失、冲突、模糊或结构不受支持时，不中止其他安全字段；保持空白或原值，并记录到核对清单。若已获准使用的方案、规范化信息表和项目知识库中仍无明确信息，则交由 CRA 人工确认或填写，不要求 Agent 推断或补填；只要状态和处理建议已完整记录，这类字段不阻断原型验收。
