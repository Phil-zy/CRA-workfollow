# 输出验证

## 通用门槛

先运行目标格式对应的验证脚本，确认输入哈希未变、输出路径和哈希与审计绑定、配置版本一致、每个字段状态和来源可追溯、核对清单逐行一致。

CRA 明确确认双输出后，运行 `validate_review_record.py` 验证人工核对审计及已核对清单，确认原输出未被覆盖、所有关联文件哈希有效、逐字段最终决定与核对人/时间/ID/总体结论一致。

## Word

运行 `scripts/validate_outputs.py` 后，优先使用 Codex `documents` Skill 提供的 `render_docx.py` 把最终 Word 渲染为每页 PNG。逐页检查无文字裁切、重叠或表格错位，字符复选框正常显示，长文本不溢出，签名及机构区域保持原样。

若缺少 LibreOffice/soffice 但存在 Microsoft Word，可用 Word 以只读方式导出临时 PDF，再逐页渲染；不得保存回输入文件。若无法完成视觉渲染，只能报告结构检查结果，不得声称完整验收通过。

## Excel

运行 `scripts/validate_xlsx_outputs.py` 后，对最终 Excel 的每个工作表至少渲染一次并检查：填写值清晰可见，无裁切或异常换行，工作表名称和顺序不变，合并区域无错位，公式区域未被覆盖，打印区域和页面方向保持不变，人工保留区域未改变。

优先使用 Codex `spreadsheets` Skill 的导入、检查和渲染能力。若最终环境无法渲染 Excel，只能报告程序化检查结果，不得声称完整验收通过。
