# CRA Skill 本地 Python 运行环境

## 结论与根因

首版 21 项测试由 Windows Store Python 3.13 启动。测试中的子进程统一使用 `sys.executable`，因此继承了该解释器的用户 `site-packages`：

- 解释器：`C:\Users\ASUS\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.13_qbz5n2kfra8p0\python.exe`
- `jsonschema==4.26.0` 来源：该解释器的用户包目录；安装记录显示它同时被本机 `mcp` 使用。
- 佐证：旧测试缓存为 `test_*cpython-313.pyc`，仓库测试代码从首版起即用 `sys.executable` 启动脚本。

实际 Skill 调用使用 Codex 捆绑 Python 3.12.13：

`C:\Users\ASUS\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`

该解释器当时已有 `openpyxl==3.1.5` 和 `python-docx==1.2.0`，但没有 `jsonschema`。仓库只有顶层 `requirements.txt`，没有项目 venv、离线依赖或统一入口，因此测试环境无法复现。

现在仓库以 `runtime/wheelhouse/`、`requirements-lock.txt`、`bootstrap_runtime.py` 和 `run.ps1` 固定环境。用户全局包不再作为依赖来源。

当前 wheelhouse 于 2026-08-17 在不含 CRA 数据的开发环境中一次性从 PyPI 获取。下载只涉及公开包名；运行与测试不再访问 PyPI。所有已下载 wheel 均由 `requirements-lock.txt` 的 SHA-256 约束，安装时缺包、增包或内容不符都会失败。

## 支持范围

- Windows x64
- CPython 3.12.x；当前验证版本为 Codex Python 3.12.13
- 项目隔离环境：`<仓库>\.runtime\cra-fill-structured-form\`
- 运行期网络策略：仅本地，pip 强制 `--no-index --no-cache-dir --require-hashes`
- 直接依赖：`jsonschema==4.26.0`、`openpyxl==3.1.5`、`python-docx==1.2.0`

`.runtime/` 是可重建的本机目录，不提交 Git。离线 wheel 和带 SHA-256 的完整锁文件随仓库提交。

## 首次部署或更换 agent

1. 使用 Codex 的工作区依赖工具取得捆绑 Python 路径。不要用 `where python` 选系统解释器。
2. 在仓库根目录执行以下命令；将示例路径替换为工具实际返回的 Python 3.12 路径：

```powershell
& "C:\Users\ASUS\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" `
  ".\.agents\skills\cra-fill-structured-form\runtime\bootstrap_runtime.py" `
  --project-root "."
```

引导器只读取仓库中的 wheelhouse。它会先在项目内唯一临时目录创建 venv，校验完整依赖集、模块来源、manifest、锁文件和 wheelhouse 精确清单，成功后才以最后一次排他重命名发布到 `.runtime\cra-fill-structured-form`；失败时不保留不完整环境，不会撤回或删除其他进程已经发布的环境，也不会读取任何 CRA 输入。

如果目标目录已存在但不完整，引导器会停止并要求人工移走该目录，不会自动删除未知文件。

## 唯一启动方式

所有业务脚本和测试均从仓库根目录通过同一个入口运行：

```powershell
# 环境与版本预检
& ".\.agents\skills\cra-fill-structured-form\run.ps1" preflight

# 模板检查示例
& ".\.agents\skills\cra-fill-structured-form\run.ps1" inspect_template.py --help

# 全部自动化回归
& ".\.agents\skills\cra-fill-structured-form\run.ps1" tests

# 仅环境级回归
& ".\.agents\skills\cra-fill-structured-form\run.ps1" tests tests.test_runtime_environment -v
```

不要直接运行 `python scripts\*.py` 或 `python -m unittest`。启动器拒绝项目外 RuntimeDir，清除 `PYTHONPATH`、`PYTHONHOME` 等外部 Python 路径，使用 `-E` 忽略 Python 环境变量；随后确认解释器位于项目内、关闭用户 site-packages，并逐项检查 10 个锁定依赖、三个关键导入的模块来源、`runtime.json`、锁文件 SHA-256 及 wheelhouse 文件与哈希，全部通过后才进入业务入口。

缺少环境或依赖时，启动器会输出需要使用的引导脚本、项目路径和运行环境路径。不得通过临时 `PYTHONPATH`、全局 `pip install` 或跳过 Schema 校验来继续。

## 离线部署和依赖更新

正常部署不需要联网：复制或检出完整仓库后，按“首次部署”执行即可。

只有维护者明确批准依赖升级时，才可在不含 CRA 数据的开发环境中一次性获取新 wheel。更新时必须同时：

1. 固定所有直接与传递依赖版本。
2. 只接受 Windows x64 / CPython 3.12 兼容 wheel，不使用源码构建。
3. 更新 `requirements-lock.txt` 中每个 wheel 的 SHA-256。
4. 从空 `.runtime` 离线引导并运行环境测试及完整回归。
5. 确认运行期仍使用 `--no-index`，且未读取、修改或上传 CRA 输入。
