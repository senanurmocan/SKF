param(
    [ValidateSet('3.11', '3.12', '3.13', '3.14')]
    [string]$PythonVersion = '3.14',
    [string]$Wheelhouse = ''
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$requirementsPath = Join-Path $projectRoot '06-Haziran\requirements.txt'
$pythonPath = Join-Path $projectRoot '.venv\Scripts\python.exe'

if (-not (Test-Path -LiteralPath $requirementsPath -PathType Leaf)) {
    throw 'The pinned application requirements file is missing.'
}

if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    & py ("-{0}" -f $PythonVersion) -m venv (Join-Path $projectRoot '.venv')
    if ($LASTEXITCODE -ne 0) {
        throw "Python $PythonVersion is unavailable or the virtual environment could not be created."
    }
}

$pipArguments = @(
    '-m', 'pip', 'install', '--disable-pip-version-check',
    '-r', $requirementsPath
)
if (-not [string]::IsNullOrWhiteSpace($Wheelhouse)) {
    if (-not (Test-Path -LiteralPath $Wheelhouse -PathType Container)) {
        throw 'The supplied approved wheelhouse directory does not exist.'
    }
    $pipArguments += @('--no-index', '--find-links', $Wheelhouse)
}

& $pythonPath @pipArguments
if ($LASTEXITCODE -ne 0) {
    throw 'Dependency installation failed. Review the approved package source or wheelhouse, then retry.'
}

Write-Host 'SKF internal server environment is ready.'
