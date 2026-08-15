from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from common import (
    ALLOWED_STATUSES,
    display_value,
    ensure_output_isolated,
    ensure_within,
    inspect_xlsx,
    load_facts,
    load_json,
    normalize_text,
    publish_new_file,
    resolved,
    sha256_file,
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


def workbook_cell(workbook, field: dict):
    target = field["target"]
    try:
        sheet = workbook[target["sheet"]]
        cell = sheet[target["cell"]]
    except (KeyError, TypeError, ValueError) as exc:
        raise ValueError(f"Excel 目标位置无效: {target}") from exc
    if isinstance(cell, MergedCell):
        raise ValueError(f"Excel 目标必须指向合并区域左上角单元格: {target}")
    return cell


def process_fields(workbook, config: dict, facts: dict[str, dict], write: bool) -> list[dict]:
    rows: list[dict] = []
    for field in config["fields"]:
        cell = workbook_cell(workbook, field)
        original = display_value(cell.value)
        mapping_type = field["mapping"].get("type")
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
        if mapping_type != "direct" or field["mapping"].get("mode") != "cell":
            rows.append(result_row(field, "待确认", fact, original, "当前 Excel 首版不支持该映射类型", "人工处理"))
            continue
        if cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("=")):
            rows.append(result_row(field, "待确认", fact, original, "目标单元格包含公式，不允许覆盖", "人工确认目标位置"))
            continue
        value = fact["value"]
        if not normalize_text(original):
            if write:
                cell.value = value
            rows.append(result_row(field, "已填写", fact, original, "已写入确认事实", "审核生成结果"))
        elif normalize_text(original) == normalize_text(value):
            rows.append(result_row(field, "已核对", fact, original, "目标已有值且与来源一致", "无需操作"))
        else:
            rows.append(result_row(field, "冲突", fact, original, "目标已有值与来源不一致", "核对冲突并决定是否生成新副本"))
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description="使用已启用配置填写 CRA Excel 表单副本")
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

    config = load_json(config_path)
    validate_config(config, "enabled")
    if config["template"]["format"] != "xlsx":
        raise ValueError("模板配置格式不是 xlsx")

    before_target_hash = sha256_file(target)
    before_facts_hash = sha256_file(facts_path)
    inspection = inspect_xlsx(target)
    if inspection["structure_fingerprint"] != config["template"]["structure_fingerprint"]:
        raise ValueError("模板结构指纹与已启用配置不匹配")
    if inspection["sha256"] != config["template"]["sha256"]:
        raise ValueError("模板 SHA-256 与已启用配置不匹配")

    workbook = load_workbook(target, data_only=False, read_only=False, keep_links=True)
    try:
        rows = process_fields(workbook, config, load_facts(facts_path), write=True)
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        output_path = unique_output_path(output_dir, f"{target.stem}_待核对_{timestamp}.xlsx")
        if sha256_file(target) != before_target_hash or sha256_file(facts_path) != before_facts_hash:
            raise RuntimeError("检测到输入文件在运行期间发生变化，停止交付")
        temporary_path = temporary_output_path(output_path)
        output_published = False
        try:
            workbook.save(temporary_path)
            if sha256_file(target) != before_target_hash or sha256_file(facts_path) != before_facts_hash:
                raise RuntimeError("检测到输入文件在运行期间发生变化，停止交付")
            publish_new_file(temporary_path, output_path)
            output_published = True
        finally:
            if temporary_path.exists():
                temporary_path.unlink()
    finally:
        workbook.close()

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
            "output_form": str(output_path),
            "output_form_sha256": sha256_file(output_path),
            "output_format": "xlsx",
            "fields": rows,
        }
        write_json_new(audit_path, audit)
    except Exception:
        if output_published and output_path.exists():
            output_path.unlink()
        raise
    print(json.dumps({"status": "completed", "output_xlsx": str(output_path), "audit_json": str(audit_path), "status_counts": {status: sum(1 for row in rows if row["status"] == status) for status in sorted(ALLOWED_STATUSES)}}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
