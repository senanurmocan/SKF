param(
    [ValidateRange(1024, 65535)]
    [int]$Port = 8505
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$pythonPath = Join-Path $projectRoot '.venv\Scripts\python.exe'
$launcherPath = Join-Path $projectRoot '06-Haziran\run.py'

if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    throw 'The project virtual environment is missing. Install the approved requirements before starting the service.'
}
if (-not (Test-Path -LiteralPath $launcherPath -PathType Leaf)) {
    throw 'The application launcher is missing from the deployment folder.'
}
if ([string]::IsNullOrWhiteSpace($env:SKF_DATA_ROOT)) {
    throw 'SKF_DATA_ROOT is not configured for this Windows service.'
}
if ([string]::IsNullOrWhiteSpace($env:SKF_SYSTEM_USAGE_ROOT)) {
    throw 'SKF_SYSTEM_USAGE_ROOT is not configured for this Windows service.'
}
if ([string]::IsNullOrWhiteSpace($env:SKF_PUBLIC_URL)) {
    throw 'SKF_PUBLIC_URL must match the HTTPS URL configured in IIS.'
}
$publicUri = $null
if (-not [uri]::TryCreate($env:SKF_PUBLIC_URL, [System.UriKind]::Absolute, [ref]$publicUri) -or $publicUri.Scheme -ne 'https') {
    throw 'SKF_PUBLIC_URL must be an absolute HTTPS URL.'
}
if ($publicUri.AbsolutePath -notin @('', '/') -or $publicUri.Query -or $publicUri.Fragment) {
    throw 'SKF_PUBLIC_URL must be the HTTPS site origin without a path, query, or fragment.'
}
if (-not (Test-Path -LiteralPath $env:SKF_DATA_ROOT -PathType Container)) {
    throw 'The service account cannot read SKF_DATA_ROOT. Check its share and NTFS permissions.'
}
if (-not (Test-Path -LiteralPath $env:SKF_SYSTEM_USAGE_ROOT -PathType Container)) {
    throw 'The service account cannot read SKF_SYSTEM_USAGE_ROOT. Check its share and NTFS permissions.'
}

$env:SKF_INTERNAL_SERVER = '1'
$env:PYTHONDONTWRITEBYTECODE = '1'
Set-Location -LiteralPath (Join-Path $projectRoot '06-Haziran')
& $pythonPath $launcherPath --server --port $Port
exit $LASTEXITCODE
