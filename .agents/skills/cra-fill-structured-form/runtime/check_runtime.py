from __future__ import annotations

import argparse
import hashlib
import importlib
import importlib.metadata
import json
import platform
import re
import site
import sys
from pathlib import Path


LOCKED_DISTRIBUTIONS = {
    "attrs": "26.1.0",
    "et-xmlfile": "2.0.0",
    "jsonschema": "4.26.0",
    "jsonschema-specifications": "2025.9.1",
    "lxml": "6.1.1",
    "openpyxl": "3.1.5",
    "python-docx": "1.2.0",
    "referencing": "0.37.0",
    "rpds-py": "2026.6.3",
    "typing-extensions": "4.16.0",
}
REQUIRED_IMPORTS = {
    "jsonschema": ("jsonschema.validators", "Draft202012Validator"),
    "openpyxl": ("openpyxl", "load_workbook"),
    "python-docx": ("docx", "Document"),
}
ALLOWED_UNTRACKED_DISTRIBUTIONS = {"pip"}


def normalize_distribution(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).lower()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def wheelhouse_inventory(wheelhouse: Path) -> dict[str, str]:
    return {
        path.name: sha256(path)
        for path in sorted(wheelhouse.glob("*.whl"), key=lambda item: item.name.lower())
    }


def validate_wheelhouse(wheelhouse: Path) -> list[str]:
    problems: list[str] = []
    wheels = sorted(wheelhouse.glob("*.whl"), key=lambda item: item.name.lower())
    found: dict[str, str] = {}
    for wheel in wheels:
        parts = wheel.name.removesuffix(".whl").split("-")
        if len(parts) < 5:
            problems.append(f"wheelhouse 含无效文件名：{wheel.name}")
            continue
        distribution = normalize_distribution(parts[0])
        version = parts[1]
        if distribution in found:
            problems.append(f"wheelhouse 含重复依赖：{distribution}")
        found[distribution] = version

    if found != LOCKED_DISTRIBUTIONS:
        missing = sorted(set(LOCKED_DISTRIBUTIONS) - set(found))
        extra = sorted(set(found) - set(LOCKED_DISTRIBUTIONS))
        mismatched = sorted(
            name
            for name in set(found) & set(LOCKED_DISTRIBUTIONS)
            if found[name] != LOCKED_DISTRIBUTIONS[name]
        )
        if missing:
            problems.append(f"wheelhouse 缺少锁定依赖：{', '.join(missing)}")
        if extra:
            problems.append(f"wheelhouse 含未锁定依赖：{', '.join(extra)}")
        for name in mismatched:
            problems.append(
                f"wheelhouse 版本不匹配：{name} 需要 {LOCKED_DISTRIBUTIONS[name]}，实际 {found[name]}"
            )
    return problems


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="检查 CRA Skill 的隔离 Python 运行环境。")
    parser.add_argument("--project-root", type=Path, required=True)
    parser.add_argument("--runtime-dir", type=Path, required=True)
    parser.add_argument(
        "--manifest-runtime-dir",
        type=Path,
        help="仅供离线引导器在原子发布前验证最终 manifest 路径。",
    )
    parser.add_argument("--json", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = args.project_root.resolve()
    runtime_dir = args.runtime_dir.resolve()
    manifest_runtime_dir = (
        args.manifest_runtime_dir.resolve() if args.manifest_runtime_dir else runtime_dir
    )
    runtime_tools = Path(__file__).resolve().parent
    lock_file = runtime_tools / "requirements-lock.txt"
    wheelhouse = runtime_tools / "wheelhouse"
    problems: list[str] = []
    versions: dict[str, str] = {}

    for checked_path, label in (
        (runtime_dir, "运行环境目录"),
        (manifest_runtime_dir, "manifest 运行环境目录"),
    ):
        try:
            relative = checked_path.relative_to(project_root)
            if not relative.parts:
                raise ValueError
        except ValueError:
            problems.append(f"{label}必须位于项目目录内：{checked_path}")

    if sys.version_info[:2] != (3, 12):
        problems.append(f"Python 版本应为 3.12.x，实际为 {platform.python_version()}")
    if Path(sys.prefix).resolve() != runtime_dir:
        problems.append(f"当前解释器不属于指定运行环境：{Path(sys.executable).resolve()}")
    if site.ENABLE_USER_SITE is not False:
        problems.append("当前解释器仍可能读取用户 site-packages")

    for distribution, expected_version in LOCKED_DISTRIBUTIONS.items():
        try:
            actual_version = importlib.metadata.version(distribution)
        except importlib.metadata.PackageNotFoundError:
            problems.append(f"缺少 {distribution}=={expected_version}")
            continue
        versions[distribution] = actual_version
        if actual_version != expected_version:
            problems.append(
                f"{distribution} 版本不匹配：需要 {expected_version}，实际 {actual_version}"
            )

    installed = {
        normalize_distribution(distribution.metadata["Name"])
        for distribution in importlib.metadata.distributions()
        if distribution.metadata.get("Name")
    }
    unexpected = sorted(installed - set(LOCKED_DISTRIBUTIONS) - ALLOWED_UNTRACKED_DISTRIBUTIONS)
    if unexpected:
        problems.append(f"运行环境含未锁定依赖：{', '.join(unexpected)}")

    for distribution, (module_name, symbol_name) in REQUIRED_IMPORTS.items():
        if distribution not in versions:
            continue
        try:
            module = importlib.import_module(module_name)
            getattr(module, symbol_name)
            module_path = Path(module.__file__).resolve()
            module_path.relative_to(runtime_dir)
        except Exception as exc:
            problems.append(f"{distribution} 无法从项目运行环境正常导入：{exc}")

    problems.extend(validate_wheelhouse(wheelhouse))
    current_lock_sha256 = sha256(lock_file) if lock_file.is_file() else ""
    current_wheelhouse = wheelhouse_inventory(wheelhouse) if wheelhouse.is_dir() else {}
    manifest_path = runtime_dir / "runtime.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        manifest = {}
        problems.append(f"无法读取运行环境 manifest：{manifest_path}：{exc}")

    expected_python = manifest_runtime_dir / "Scripts" / "python.exe"
    manifest_checks = {
        "schema_version": 1,
        "runtime_dir": str(manifest_runtime_dir),
        "python_executable": str(expected_python),
        "requirements_lock_sha256": current_lock_sha256,
        "network_policy": "offline-only",
        "dependencies": versions,
        "wheelhouse": current_wheelhouse,
    }
    for key, expected in manifest_checks.items():
        if manifest.get(key) != expected:
            if key == "requirements_lock_sha256":
                problems.append("runtime.json 与当前 requirements-lock.txt 不一致")
            elif key == "wheelhouse":
                problems.append("runtime.json 与当前 wheelhouse 文件及哈希不一致")
            else:
                problems.append(f"runtime.json 字段不匹配：{key}")

    if problems:
        bootstrap = runtime_tools / "bootstrap_runtime.py"
        print("CRA Skill Python 运行环境检查失败：", file=sys.stderr)
        for problem in problems:
            print(f"- {problem}", file=sys.stderr)
        print("请使用 Codex 工作区依赖工具返回的 Python 3.12 执行：", file=sys.stderr)
        print(
            f'& "<Codex Python 3.12 路径>" "{bootstrap}" '
            f'--project-root "{project_root}" --runtime-dir "{manifest_runtime_dir}"',
            file=sys.stderr,
        )
        print("引导过程只读取仓库 wheelhouse，不访问网络，也不会读取 CRA 输入。", file=sys.stderr)
        return 2

    result = {
        "status": "passed",
        "python_executable": str(Path(sys.executable).resolve()),
        "python_version": platform.python_version(),
        "runtime_dir": str(runtime_dir),
        "user_site_enabled": site.ENABLE_USER_SITE,
        "dependencies": versions,
        "network_policy": "offline-only",
        "requirements_lock_sha256": current_lock_sha256,
        "wheelhouse": current_wheelhouse,
    }
    if args.json:
        print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    else:
        print(
            "CRA Skill Python 运行环境检查通过："
            f"Python {result['python_version']}，"
            f"jsonschema {versions['jsonschema']}，"
            f"openpyxl {versions['openpyxl']}，"
            f"python-docx {versions['python-docx']}。"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
