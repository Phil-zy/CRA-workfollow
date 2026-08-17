[CmdletBinding(PositionalBinding = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet(
        'preflight',
        'tests',
        'activate_mapping.py',
        'build_checklist.py',
        'draft_mapping.py',
        'fill_docx.py',
        'fill_xlsx.py',
        'finalize_review.py',
        'inspect_template.py',
        'validate_outputs.py',
        'validate_review_record.py',
        'validate_xlsx_outputs.py'
    )]
    [string] $EntryPoint,

    [Parameter()]
    [string] $RuntimeDir,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $EntryPointArgs
)

$ErrorActionPreference = 'Stop'
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
if ([string]::IsNullOrWhiteSpace($RuntimeDir)) {
    $RuntimeDir = Join-Path $projectRoot '.runtime\cra-fill-structured-form'
}
$RuntimeDir = [System.IO.Path]::GetFullPath($RuntimeDir)
$projectPrefix = $projectRoot.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
if (-not $RuntimeDir.StartsWith($projectPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    [Console]::Error.WriteLine("RuntimeDir must be inside the project: $projectRoot")
    exit 2
}
$python = Join-Path $RuntimeDir 'Scripts\python.exe'
$runtimeTools = Join-Path $PSScriptRoot 'runtime'
$preflight = Join-Path $runtimeTools 'check_runtime.py'
$bootstrap = Join-Path $runtimeTools 'bootstrap_runtime.py'

if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    [Console]::Error.WriteLine('CRA Skill local Python runtime is missing or incomplete.')
    [Console]::Error.WriteLine('Run the following command with the Codex workspace Python 3.12:')
    [Console]::Error.WriteLine(
        ('& "<Codex Python 3.12 路径>" "{0}" --project-root "{1}" --runtime-dir "{2}"' -f $bootstrap, $projectRoot, $RuntimeDir)
    )
    [Console]::Error.WriteLine('Bootstrap uses only the repository wheelhouse; it does not access the network or CRA inputs.')
    exit 2
}

Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONHOME -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONSTARTUP -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONUSERBASE -ErrorAction SilentlyContinue
$env:PYTHONNOUSERSITE = '1'
$env:PIP_NO_INDEX = '1'
$env:PIP_DISABLE_PIP_VERSION_CHECK = '1'
$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONUTF8 = '1'

& $python -E -X utf8 $preflight --project-root $projectRoot --runtime-dir $RuntimeDir --json | Out-Null
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

if ($EntryPoint -eq 'preflight') {
    & $python -E -X utf8 $preflight --project-root $projectRoot --runtime-dir $RuntimeDir
    exit $LASTEXITCODE
}

$previousLocation = Get-Location
try {
    Set-Location -LiteralPath $projectRoot
    if ($EntryPoint -eq 'tests') {
        if ($EntryPointArgs.Count -eq 0) {
            & $python -E -X utf8 -m unittest discover -s tests -v
        }
        else {
            & $python -E -X utf8 -m unittest @EntryPointArgs
        }
    }
    else {
        $script = Join-Path (Join-Path $PSScriptRoot 'scripts') $EntryPoint
        & $python -E -X utf8 $script @EntryPointArgs
    }
    exit $LASTEXITCODE
}
finally {
    Set-Location -LiteralPath $previousLocation
}
