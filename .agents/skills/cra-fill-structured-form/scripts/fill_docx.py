from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from docx import Document

from common import (
    ALLOWED_STATUSES,
    compact_trailing_empty_paragraph,
    ensure_output_isolated,
    ensure_within,
    inspect_docx,
    load_facts,
    load_json,
    normalize_text,
    publish_new_file,
    replace_choice,
    resolved,
    selected_options,
    set_blank_cell_text,
    sha256_file,
    target_cell,
    target_location,
    temporary_output_path,
    unique_output_path,
    validate_config,
    write_json_new,
)


def result_row(field: dict, status: str, fact: dict | None, original: str, reason: str, action: str) -> dict:
    if status not in ALLOWED_STATUSES:
        raise ValueError(f"非法处理状态: {status}")
    return {
        "field_id": field["id"],
        "field_name": field["label"],
        "target_location": target_location(field),
        "status": status,
        "source_value": "" if fact is None else fact["value"],
        "target_original": original,
        "reason": reason,
        "source_file": "" if fact is None else fact["source_file"],
        "source_location": "" if fact is None else fact["source_location"],
        "suggested_action": action,
        "cra_final_decision": "",
    }


def process_direct(document: Document, field: dict, fact: dict, write: bool) -> dict:
    cell = target_cell(document, field["target"])
    original = cell.text
    mapping = field["mapping"]
    value = fact["value"]
    if mapping.get("mode") == "inline_after_anchor":
        anchor = mapping["anchor"]
        if anchor not in original:
            return result_row(field, "待确认", fact, original, "未找到配置的行内锚点", "人工确认目标位置")
        existing = original.split(anchor, 1)[1].strip()
        if existing:
            if normalize_text(existing) == normalize_text(value):
                return result_row(field, "已核对", fact, original, "目标已有值且与来源一致", "无需操作")
            return result_row(field, "冲突", fact, original, "目标已有值与来源不一致", "核对冲突并决定是否生成新副本")
        if not write:
            return result_row(field, "已填写", fact, original, "可安全写入", "审核生成结果")
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if anchor in run.text:
                    run.text = run.text.replace(anchor, anchor + value, 1)
                    return result_row(field, "已填写", fact, original, "已写入确认事实", "审核生成结果")
        return result_row(field, "待确认", fact, original, "锚点跨越多个文本 run，无法保留格式写入", "人工填写")

    if not normalize_text(original):
        if write and not set_blank_cell_text(cell, value):
            return result_row(field, "待确认", fact, original, "目标不再为空", "重新检查目标")
        return result_row(field, "已填写", fact, original, "已写入确认事实", "审核生成结果")
    if normalize_text(original) == normalize_text(value):
        return result_row(field, "已核对", fact, original, "目标已有值且与来源一致", "无需操作")
    return result_row(field, "冲突", fact, original, "目标已有值与来源不一致", "核对冲突并决定是否生成新副本")


def process_choice(document: Document, field: dict, fact: dict, write: bool) -> dict:
    cell = target_cell(document, field["target"])
    original = cell.text
    mapping = field["mapping"]
    choice = mapping.get("value_to_option", {}).get(normalize_text(fact["value"]))
    if not isinstance(choice, str) or not choice:
        return result_row(field, "待确认", fact, original, "来源事实不能唯一匹配模板选项", "CRA 确认选项；不要推断")
    options = list(mapping["value_to_option"].values())
    selected = selected_options(original, options, mapping["checked"])
    if len(selected) == 1 and selected[0] == choice:
        return result_row(field, "已核对", fact, original, "目标选项已勾选且与来源一致", "无需操作")
    if selected:
        return result_row(field, "冲突", fact, original, "目标已有其他选项被勾选", "核对冲突并决定是否生成新副本")
    if not write:
        return result_row(field, "已填写", fact, original, "选项可唯一匹配", "审核生成结果")
    if replace_choice(cell, choice, mapping["unchecked"], mapping["checked"]):
        return result_row(field, "已填写", fact, original, "已勾选唯一匹配选项", "审核生成结果")
    return result_row(field, "待确认", fact, original, "选项符号和文字无法跨文本 run 唯一定位", "人工填写并检查格式")


def process_fields(document: Document, config: dict, facts: dict[str, dict], write: bool) -> list[dict]:
    rows = []
    for field in config["fields"]:
        mapping_type = field["mapping"].get("type")
        cell = target_cell(document, field["target"])
        original = cell.text
        if mapping_type == "manual":
            rows.append(result_row(field, "人工保留", None, original, "签名或机构专用区域", "保持原样并由责任人填写"))
            continue
        fact = facts.get(field.get("fact_key"))
        if fact is None or fact["status"] == "missing":
            rows.append(result_row(field, "缺失", fact, original, "没有可用的已确认事实", "由 CRA 人工确认或填写"))
            continue
        if fact["status"] != "confirmed":
            rows.append(result_row(field, "待确认", fact, original, "来源事实未确认", "由 CRA 人工确认或填写"))
            continue
        if mapping_type == "direct":
            rows.append(process_direct(document, field, fact, write=write))
        elif mapping_type == "choice":
            rows.append(process_choice(document, field, fact, write=write))
        else:
            rows.append(result_row(field, "待确认", fact, original, "不支持的映射类型", "人工处理"))
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description="使用已启用配置填写 CRA Word 表单副本")
    parser.add_argument("--target", required=True)
    parser.add_argument("--facts", required=True)
    parser.add_argument("--config", required=True)
    parser.add_argument("--authorized-workspace", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--audit-json", required=True)
    parser.add_argument("--authorization-id", required=True)
    args = parser.parse_args()

    workspace = resolved(args.authorized_workspace)
    target = resolved(args.target)
    facts_path = resolved(args.facts)
    config_path = resolved(args.config)
    output_dir = resolved(args.output_dir)
    audit_path = resolved(args.audit_json)
    ensure_within(workspace, [target, facts_path, config_path, output_dir, audit_path])
    ensure_output_isolated(output_dir, [target, facts_path, config_path])
    ensure_output_isolated(audit_path.parent, [target, facts_path, config_path])
    if not args.authorization_id.strip():
        raise ValueError("缺少明确的授权确认 ID")
    if audit_path.exists():
        raise FileExistsError(f"审计文件已存在，不得覆盖: {audit_path}")
    if output_dir in (target, facts_path) or audit_path in (target, facts_path):
        raise ValueError("输出路径不得与输入路径相同")

    config = load_json(config_path)
    validate_config(config, "enabled")
    approval = config["approval"]
    if not all(approval.get(key) for key in ("approved_by", "approved_at", "approval_id")):
        raise ValueError("已启用配置缺少完整 CRA 审批记录")
    if any(field.get("mapping_review") != "approved" for field in config["fields"]):
        raise ValueError("存在未批准的字段映射")

    before_target_hash = sha256_file(target)
    before_facts_hash = sha256_file(facts_path)
    inspection = inspect_docx(target)
    if inspection["structure_fingerprint"] != config["template"]["structure_fingerprint"]:
        raise ValueError("模板结构指纹与已启用配置不匹配")
    if inspection["sha256"] != config["template"]["sha256"]:
        raise ValueError("模板 SHA-256 与已启用配置不匹配")

    facts = load_facts(facts_path)
    document = Document(target)
    rows = process_fields(document, config, facts, write=True)

    compact_trailing_empty_paragraph(document)

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_name = f"{target.stem}_待核对_{timestamp}.docx"
    output_path = unique_output_path(output_dir, output_name)
    if output_path.resolve(strict=False) == target.resolve(strict=True):
        raise ValueError("输出文件不得覆盖输入文件")
    if sha256_file(target) != before_target_hash or sha256_file(facts_path) != before_facts_hash:
        raise RuntimeError("检测到输入文件在运行期间发生变化，停止交付")
    temporary_path = temporary_output_path(output_path)
    output_published = False
    try:
        document.save(temporary_path)
        if sha256_file(target) != before_target_hash or sha256_file(facts_path) != before_facts_hash:
            raise RuntimeError("检测到输入文件在运行期间发生变化，停止交付")
        publish_new_file(temporary_path, output_path)
        output_published = True
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    try:
        audit = {
            "schema_version": "1.0",
            "executed_at": datetime.now(timezone.utc).isoformat(),
            "authorization_id": args.authorization_id,
            "network_used": False,
            "external_services_used": False,
            "target_input": str(target),
            "target_input_sha256": before_target_hash,
            "facts_input": str(facts_path),
            "facts_input_sha256": before_facts_hash,
            "config": str(config_path),
            "config_sha256": sha256_file(config_path),
            "config_id": config["config_id"],
            "config_version": config["version"],
            "output_docx": str(output_path),
            "output_docx_sha256": sha256_file(output_path),
            "output_form": str(output_path),
            "output_form_sha256": sha256_file(output_path),
            "output_format": "docx",
            "fields": rows,
        }
        write_json_new(audit_path, audit)
    except Exception:
        if output_published and output_path.exists():
            output_path.unlink()
        raise
    print(json.dumps({"status": "completed", "output_docx": str(output_path), "audit_json": str(audit_path), "status_counts": {status: sum(1 for row in rows if row["status"] == status) for status in sorted(ALLOWED_STATUSES)}}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
