from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone

from common import ensure_within, load_json, resolved, validate_config, write_json_new


def main() -> None:
    parser = argparse.ArgumentParser(description="经 CRA 批准后把模板配置草稿复制为不可变启用版本")
    parser.add_argument("--draft", required=True)
    parser.add_argument("--authorized-workspace", required=True)
    parser.add_argument("--config-library", required=True)
    parser.add_argument("--approved-by", required=True)
    parser.add_argument("--approval-id", required=True)
    args = parser.parse_args()

    workspace = resolved(args.authorized_workspace)
    draft_path = resolved(args.draft)
    config_library = resolved(args.config_library)
    ensure_within(workspace, [draft_path, config_library])
    config = load_json(draft_path)
    validate_config(config, "draft")
    if draft_path.parent != config_library / "drafts":
        raise ValueError("草稿必须位于配置库 drafts 目录")

    activated = json.loads(json.dumps(config, ensure_ascii=False))
    activated["status"] = "enabled"
    activated["enabled_at"] = datetime.now(timezone.utc).isoformat()
    activated["approval"] = {"approved_by": args.approved_by, "approved_at": activated["enabled_at"], "approval_id": args.approval_id}
    for field in activated["fields"]:
        field["mapping_review"] = "approved"
    validate_config(activated, "enabled")

    output = config_library / "enabled" / draft_path.name
    write_json_new(output, activated)
    print(json.dumps({"status": "enabled", "config": str(output)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
