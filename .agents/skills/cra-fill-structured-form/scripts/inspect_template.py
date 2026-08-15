from __future__ import annotations

import argparse
import json
from pathlib import Path

from common import ensure_within, inspect_template, matching_configs, resolved, write_json_new


def main() -> None:
    parser = argparse.ArgumentParser(description="只读检查 CRA Word 或 Excel 模板及配置匹配状态")
    parser.add_argument("--target", required=True)
    parser.add_argument("--authorized-workspace", required=True)
    parser.add_argument("--config-library", required=True)
    parser.add_argument("--output")
    args = parser.parse_args()

    workspace = resolved(args.authorized_workspace)
    target = resolved(args.target)
    config_library = resolved(args.config_library)
    ensure_within(workspace, [target, config_library])

    inspection = inspect_template(target)
    result = {
        "target": str(target),
        "template": {key: inspection[key] for key in ("file_name", "sha256", "structure_fingerprint", "format")},
        "matching_enabled_configs": [str(path) for path in matching_configs(config_library, inspection["structure_fingerprint"], inspection["sha256"], "enabled")],
        "matching_draft_configs": [str(path) for path in matching_configs(config_library, inspection["structure_fingerprint"], inspection["sha256"], "draft")],
    }
    if inspection["format"] == "docx":
        result.update(
            paragraph_count=inspection["paragraph_count"],
            section_count=inspection["section_count"],
            table_count=inspection["table_count"],
        )
    else:
        result.update(sheet_count=inspection["sheet_count"], sheet_names=inspection["sheet_names"])
    if args.output:
        output = resolved(args.output)
        ensure_within(workspace, [output])
        write_json_new(output, result)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
