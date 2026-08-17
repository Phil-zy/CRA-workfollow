from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
import uuid
import venv
from datetime import datetime, timezone
from pathlib import Path

from check_runtime import LOCKED_DISTRIBUTIONS, sha256, validate_wheelhouse, wheelhouse_inventory


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从仓库 wheelhouse 离线创建 CRA Skill 的隔离 Python 环境。"
    )
    parser.add_argument("--project-root", type=Path, required=True)
    parser.add_argument("--runtime-dir", type=Path)
    return parser.parse_args()


def python_in(runtime_dir: Path) -> Path:
    return runtime_dir / "Scripts" / "python.exe"


def clean_environment() -> dict[str, str]:
    environment = os.environ.copy()
    for name in (
        "PYTHONHOME",
        "PYTHONPATH",
        "PIP_INDEX_URL",
        "PIP_EXTRA_INDEX_URL",
        "PIP_TRUSTED_HOST",
    ):
        environment.pop(name, None)
    environment.update(
        {
            "PIP_NO_INDEX": "1",
            "PIP_DISABLE_PIP_VERSION_CHECK": "1",
            "PYTHONNOUSERSITE": "1",
            "PYTHONIOENCODING": "utf-8",
        }
    )
    return environment


def run_checked(command: list[str], environment: dict[str, str]) -> subprocess.CompletedProcess[str]:
    process = subprocess.run(
        command,
        env=environment,
        text=True,
        encoding="utf-8",
        capture_output=True,
        check=False,
    )
    if process.returncode != 0:
        raise RuntimeError(
            f"命令失败（{process.returncode}）：{' '.join(command)}\n"
            f"STDOUT:\n{process.stdout}\nSTDERR:\n{process.stderr}"
        )
    return process


def ensure_inside_project(path: Path, project_root: Path) -> None:
    try:
        relative = path.relative_to(project_root)
        if not relative.parts:
            raise ValueError
    except ValueError as exc:
        raise ValueError(f"运行环境目录必须位于项目内：{path}") from exc


def main() -> int:
    args = parse_args()
    project_root = args.project_root.resolve()
    runtime_dir = (
        args.runtime_dir.resolve()
        if args.runtime_dir
        else project_root / ".runtime" / "cra-fill-structured-form"
    )
    ensure_inside_project(runtime_dir, project_root)

    if os.name != "nt" or platform.machine().lower() not in {"amd64", "x86_64"}:
        raise SystemExit("当前离线包仅支持 Windows x64。")
    if sys.version_info[:2] != (3, 12):
        raise SystemExit(
            f"必须使用 Codex Python 3.12 创建运行环境；当前为 {platform.python_version()}："
            f"{Path(sys.executable).resolve()}"
        )

    skill_root = Path(__file__).resolve().parents[1]
    wheelhouse = Path(__file__).resolve().parent / "wheelhouse"
    lock_file = Path(__file__).resolve().parent / "requirements-lock.txt"
    check_script = Path(__file__).resolve().parent / "check_runtime.py"
    environment = clean_environment()

    if runtime_dir.exists():
        existing_python = python_in(runtime_dir)
        if not existing_python.is_file():
            raise SystemExit(
                f"运行环境目录已存在但不完整：{runtime_dir}。"
                "请人工移走该目录后重新执行引导。"
            )
        check = run_checked(
            [
                str(existing_python),
                str(check_script),
                "--project-root",
                str(project_root),
                "--runtime-dir",
                str(runtime_dir),
                "--json",
            ],
            environment,
        )
        print(check.stdout.strip())
        return 0

    if not lock_file.is_file() or not wheelhouse.is_dir():
        raise SystemExit("仓库缺少 requirements-lock.txt 或 wheelhouse，无法离线创建环境。")
    wheelhouse_problems = validate_wheelhouse(wheelhouse)
    if wheelhouse_problems:
        raise SystemExit("wheelhouse 校验失败：\n- " + "\n- ".join(wheelhouse_problems))

    runtime_dir.parent.mkdir(parents=True, exist_ok=True)
    temporary_dir = runtime_dir.with_name(f"{runtime_dir.name}.tmp-{uuid.uuid4().hex}")
    try:
        venv.EnvBuilder(with_pip=True, clear=False, symlinks=False).create(temporary_dir)
        temporary_python = python_in(temporary_dir)
        run_checked(
            [
                str(temporary_python),
                "-m",
                "pip",
                "install",
                "--no-index",
                "--no-cache-dir",
                "--require-hashes",
                "--find-links",
                str(wheelhouse),
                "--requirement",
                str(lock_file),
            ],
            environment,
        )
        manifest = {
            "schema_version": 1,
            "created_at_utc": datetime.now(timezone.utc).isoformat(),
            "base_python": str(Path(sys.executable).resolve()),
            "base_python_version": platform.python_version(),
            "python_executable": str(python_in(runtime_dir).resolve()),
            "runtime_dir": str(runtime_dir),
            "skill_root": str(skill_root),
            "requirements_lock_sha256": sha256(lock_file),
            "network_policy": "offline-only",
            "dependencies": LOCKED_DISTRIBUTIONS,
            "wheelhouse": wheelhouse_inventory(wheelhouse),
        }
        (temporary_dir / "runtime.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        run_checked(
            [
                str(temporary_python),
                str(check_script),
                "--project-root",
                str(project_root),
                "--runtime-dir",
                str(temporary_dir),
                "--manifest-runtime-dir",
                str(runtime_dir),
                "--json",
            ],
            environment,
        )
        temporary_dir.rename(runtime_dir)
        print(json.dumps(manifest, ensure_ascii=False, sort_keys=True))
        return 0
    except Exception as exc:
        raise SystemExit(f"离线运行环境创建失败，未保留不完整环境：{exc}") from exc
    finally:
        if temporary_dir.exists():
            shutil.rmtree(temporary_dir)


if __name__ == "__main__":
    raise SystemExit(main())
