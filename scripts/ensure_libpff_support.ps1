param(
    [Parameter(Mandatory=$true)]
    [string]$PythonExe
)

$ErrorActionPreference = "Stop"

function Test-LibPffInstalled {
    param(
        [string]$Python
    )

    & $Python "-c" "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('pypff') else 1)" *> $null
    return $LASTEXITCODE -eq 0
}

if (Test-LibPffInstalled -Python $PythonExe) {
    Write-Host "[Outlook Sidecar] PST/OST support already available."
    exit 0
}

Write-Host "[Outlook Sidecar] Downloading Microsoft Visual C++ Build Tools..."
$bootstrapperPath = Join-Path $env:TEMP "vs_BuildTools.exe"

try {
    if (-not (Test-Path $bootstrapperPath)) {
        Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vs_BuildTools.exe" -OutFile $bootstrapperPath -UseBasicParsing
    }
}
catch {
    Write-Warning "Failed to download Visual Studio Build Tools. $_"
    exit 1
}

Write-Host "[Outlook Sidecar] Installing required C++ components (this can take several minutes)..."
$installArgs = @(
    "--quiet",
    "--wait",
    "--norestart",
    "--nocache",
    "--add", "Microsoft.VisualStudio.Workload.VCTools",
    "--add", "Microsoft.VisualStudio.Component.VC.Tools.x86.x64",
    "--includeRecommended",
    "--includeOptional"
)

& $bootstrapperPath @installArgs
$buildExit = $LASTEXITCODE
if ($buildExit -ne 0) {
    Write-Warning "Microsoft Visual C++ Build Tools installation returned exit code $buildExit."
    exit $buildExit
}

Write-Host "[Outlook Sidecar] Retrying libpff-python installation..."
& $PythonExe -m pip install libpff-python>=20231205
if ($LASTEXITCODE -ne 0) {
    Write-Warning "libpff-python still failed to install automatically."
    exit $LASTEXITCODE
}

if (Test-LibPffInstalled -Python $PythonExe) {
    Write-Host "[Outlook Sidecar] PST/OST support installed successfully."
    exit 0
}
else {
    Write-Warning "libpff-python installation succeeded but the module was not detected."
    exit 1
}
