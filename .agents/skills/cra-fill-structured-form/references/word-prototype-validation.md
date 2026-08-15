# Word 原型验证（兼容入口）

通用验证规则已经迁移到 [output-validation.md](output-validation.md)。本文件仅为旧流程兼容入口。

## 结构检查

运行 `scripts/validate_outputs.py`，确认输入文件 SHA-256 不变、输出可打开、段落和表格几何不变、人工保留区域不变、输出与输入路径不同、状态集合合法且核对清单完整。

## 逐页渲染

优先使用 Codex `documents` Skill 提供的 `render_docx.py` 把最终 Word 渲染为每页 PNG。若本机缺少 LibreOffice/soffice，但存在 Microsoft Word，可用 Word 以只读方式导出临时 PDF，再逐页渲染；不得保存回输入文件。

逐页以 100% 比例检查无文字裁切、重叠或表格错位，`☒` 正常显示，长文本不溢出，签名及机构区域保持原样。填写后由 1 页增加为 2 页本身不判失败；记录页数变化，并重点查找大面积字体、缩进、行距或其他格式错乱的原因。发现此类错乱时停止交付，修正后重新生成独立副本并逐页复验。

若既不能使用 LibreOffice/soffice，也不能使用本机 Word 完成只读渲染，可以完成结构检查，但必须明确说明未通过视觉渲染门槛；不得声称完整验收通过。
