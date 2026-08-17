from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / ".agents" / "skills" / "cra-fill-structured-form"
BOOTSTRAP = SKILL / "runtime" / "bootstrap_runtime.py"
RUNNER = SKILL / "run.ps1"


def clean_environment() -> dict[str, str]:
    environment = os.environ.copy()
    for name in ("PYTHONPATH", "PYTHONHOME", "PIP_INDEX_URL", "PIP_EXTRA_INDEX_URL"):
        environment.pop(name, None)
    environment.update(
        {
            "PIP_NO_INDEX": "1",
            "PYTHONNOUSERSITE": "1",
            "PYTHONIOENCODING": "utf-8",
        }
    )
    return environment


def run_checked(command: list[str], environment: dict[str, str]) -> subprocess.CompletedProcess[str]:
    process = subprocess.run(
        command,
        cwd=ROOT,
        env=environment,
        text=True,
        encoding="utf-8",
        capture_output=True,
        check=False,
    )
    if process.returncode != 0:
        raise AssertionError(
            f"command failed ({process.returncode}): {' '.join(command)}\n"
            f"STDOUT:\n{process.stdout}\nSTDERR:\n{process.stderr}"
        )
    return process


class RuntimeEnvironmentTests(unittest.TestCase):
    def test_test_process_uses_a_repository_managed_runtime(self) -> None:
        runtime_dir = Path(sys.prefix).resolve()
        self.assertTrue(runtime_dir.is_relative_to(ROOT))
        manifest_path = runtime_dir / "runtime.json"
        self.assertTrue(manifest_path.is_file())
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(Path(manifest["python_executable"]).resolve(), Path(sys.executable).resolve())
        self.assertEqual(manifest["network_policy"], "offline-only")

    def test_preflight_reports_actionable_missing_dependency_error(self) -> None:
        environment = clean_environment()
        base_python = Path(getattr(sys, "_base_executable", sys.executable)).resolve()
        check_runtime = SKILL / "runtime" / "check_runtime.py"

        with tempfile.TemporaryDirectory(dir=ROOT / ".runtime") as temporary_directory:
            naked_runtime = Path(temporary_directory) / "naked-runtime"
            run_checked(
                [str(base_python), "-m", "venv", "--without-pip", str(naked_runtime)],
                environment,
            )
            process = subprocess.run(
                [
                    str(naked_runtime / "Scripts" / "python.exe"),
                    str(check_runtime),
                    "--project-root",
                    str(ROOT),
                    "--runtime-dir",
                    str(naked_runtime),
                ],
                cwd=ROOT,
                env=environment,
                text=True,
                encoding="utf-8",
                capture_output=True,
                check=False,
            )
            self.assertEqual(process.returncode, 2)
            self.assertIn("缺少 jsonschema==4.26.0", process.stderr)
            self.assertIn("缺少 openpyxl==3.1.5", process.stderr)
            self.assertIn("缺少 python-docx==1.2.0", process.stderr)
            self.assertIn("bootstrap_runtime.py", process.stderr)
            self.assertIn("不访问网络", process.stderr)

    def test_runner_ignores_external_python_paths_and_rejects_external_runtime(self) -> None:
        environment = clean_environment()
        with tempfile.TemporaryDirectory(dir=ROOT / ".runtime") as temporary_directory:
            poison_root = Path(temporary_directory) / "poison"
            poison_package = poison_root / "jsonschema"
            poison_package.mkdir(parents=True)
            (poison_package / "__init__.py").write_text(
                "raise RuntimeError('external PYTHONPATH was imported')\n",
                encoding="utf-8",
            )
            environment["PYTHONPATH"] = str(poison_root)
            run_checked(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(RUNNER),
                    "preflight",
                ],
                environment,
            )

        external = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(RUNNER),
                "-RuntimeDir",
                str(Path(os.environ.get("TEMP", "C:/Windows/Temp")) / "external-cra-runtime"),
                "preflight",
            ],
            cwd=ROOT,
            env=clean_environment(),
            text=True,
            encoding="utf-8",
            capture_output=True,
            check=False,
        )
        self.assertEqual(external.returncode, 2)
        self.assertIn("inside the project", external.stderr)

    def test_clean_offline_runtime_loads_entrypoints_and_word_workflow(self) -> None:
        environment = clean_environment()
        base_python = Path(getattr(sys, "_base_executable", sys.executable)).resolve()

        with tempfile.TemporaryDirectory(dir=ROOT / ".runtime") as temporary_directory:
            runtime_dir = Path(temporary_directory) / "runtime"
            bootstrap = run_checked(
                [
                    str(base_python),
                    str(BOOTSTRAP),
                    "--project-root",
                    str(ROOT),
                    "--runtime-dir",
                    str(runtime_dir),
                ],
                environment,
            )
            manifest = json.loads(bootstrap.stdout)
            runtime_python = Path(manifest["python_executable"]).resolve()
            self.assertTrue(runtime_python.is_file())
            self.assertEqual(manifest["network_policy"], "offline-only")

            inspection = run_checked(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(RUNNER),
                    "-RuntimeDir",
                    str(runtime_dir),
                    "inspect_template.py",
                    "--target",
                    str(ROOT / "fixtures" / "form-fill-pilot" / "template" / "XX医院 伦理审查申请表.docx"),
                    "--authorized-workspace",
                    str(ROOT),
                    "--config-library",
                    str(ROOT / "template-configs" / "cra-fill-structured-form"),
                ],
                environment,
            )
            inspection_result = json.loads(inspection.stdout)
            self.assertTrue(inspection_result["matching_enabled_configs"])

            for entrypoint in (
                "inspect_template.py",
                "fill_docx.py",
                "build_checklist.py",
                "validate_outputs.py",
                "fill_xlsx.py",
                "validate_xlsx_outputs.py",
            ):
                run_checked(
                    [
                        "powershell.exe",
                        "-NoProfile",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-File",
                        str(RUNNER),
                        "-RuntimeDir",
                        str(runtime_dir),
                        entrypoint,
                        "--help",
                    ],
                    environment,
                )

            run_checked(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(RUNNER),
                    "-RuntimeDir",
                    str(runtime_dir),
                    "tests",
                    "tests.test_word_prototype.WordPrototypeTests.test_existing_word_sample_preserves_six_statuses_and_never_overwrites",
                ],
                environment,
            )

            isolated_tools = Path(temporary_directory) / "runtime-tools"
            shutil.copytree(SKILL / "runtime", isolated_tools)
            extra_wheel = isolated_tools / "wheelhouse" / "unexpected-0-py3-none-any.whl"
            extra_wheel.write_bytes(b"not a wheel")
            unexpected_wheel = subprocess.run(
                [
                    str(runtime_python),
                    str(isolated_tools / "check_runtime.py"),
                    "--project-root",
                    str(ROOT),
                    "--runtime-dir",
                    str(runtime_dir),
                ],
                cwd=ROOT,
                env=environment,
                text=True,
                encoding="utf-8",
                capture_output=True,
                check=False,
            )
            self.assertEqual(unexpected_wheel.returncode, 2)
            self.assertIn("wheelhouse", unexpected_wheel.stderr)

            manifest_path = runtime_dir / "runtime.json"
            stale_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            stale_manifest["requirements_lock_sha256"] = "0" * 64
            manifest_path.write_text(
                json.dumps(stale_manifest, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            stale = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(RUNNER),
                    "-RuntimeDir",
                    str(runtime_dir),
                    "preflight",
                ],
                cwd=ROOT,
                env=environment,
                text=True,
                encoding="utf-8",
                capture_output=True,
                check=False,
            )
            self.assertEqual(stale.returncode, 2)
            self.assertIn("requirements-lock.txt", stale.stderr)


if __name__ == "__main__":
    unittest.main()
