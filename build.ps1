[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$projectRoot = [System.IO.Path]::GetFullPath($PSScriptRoot)
$projectPrefix = $projectRoot.TrimEnd("\") + "\"

function Get-ProjectPath {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $fullPath = [System.IO.Path]::GetFullPath((Join-Path $projectRoot $RelativePath))
    if (-not $fullPath.StartsWith(
        $projectPrefix,
        [System.StringComparison]::OrdinalIgnoreCase
    )) {
        throw "Unsafe path outside project: $fullPath"
    }
    return $fullPath
}

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE`: $FilePath $($Arguments -join ' ')"
    }
}

function Get-BootstrapPython {
    $developmentPython = Get-ProjectPath ".venv\Scripts\python.exe"
    if (Test-Path -LiteralPath $developmentPython -PathType Leaf) {
        return @{ FilePath = $developmentPython; PrefixArguments = @() }
    }

    $localPython = Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\python.exe"
    if (Test-Path -LiteralPath $localPython -PathType Leaf) {
        return @{ FilePath = $localPython; PrefixArguments = @() }
    }

    $pyLauncher = Get-Command "py.exe" -ErrorAction SilentlyContinue
    if ($null -ne $pyLauncher) {
        return @{ FilePath = $pyLauncher.Source; PrefixArguments = @("-3.12") }
    }

    $pythonCommand = Get-Command "python.exe" -ErrorAction SilentlyContinue
    if ($null -ne $pythonCommand) {
        return @{ FilePath = $pythonCommand.Source; PrefixArguments = @() }
    }

    throw "Python 3.12 was not found. Install Python 3.12 or create the project .venv first."
}

$buildEnvironment = Get-ProjectPath ".build-venv"
$buildPython = Join-Path $buildEnvironment "Scripts\python.exe"
$requirementsFile = Get-ProjectPath "requirements-build.txt"
$versionFile = Get-ProjectPath "K_Tools\version.py"
$selfTestScript = Get-ProjectPath "K_Tools\build_self_test.py"
$specFile = Get-ProjectPath "K_Tools\K_Tools.spec"
$workDirectory = Get-ProjectPath "build\pyinstaller"
$stageDirectory = Get-ProjectPath "build\release"
$stageDist = Join-Path $stageDirectory "dist"
$pythonTestReport = Join-Path $stageDirectory "python-gis-self-test.txt"
$frozenTestReport = Join-Path $stageDirectory "frozen-gis-self-test.txt"
$distDirectory = Get-ProjectPath "dist"

if (-not (Test-Path -LiteralPath $buildPython -PathType Leaf)) {
    if (Test-Path -LiteralPath $buildEnvironment) {
        # This directory is owned solely by this script and was validated above.
        Remove-Item -LiteralPath $buildEnvironment -Recurse -Force
    }

    Write-Host "Creating isolated build environment..." -ForegroundColor Cyan
    $bootstrap = Get-BootstrapPython
    $bootstrapArguments = @($bootstrap.PrefixArguments) + @("-m", "venv", $buildEnvironment)
    Invoke-Native -FilePath $bootstrap.FilePath -Arguments $bootstrapArguments
}

Write-Host "Synchronizing pinned build dependencies..." -ForegroundColor Cyan
Invoke-Native -FilePath $buildPython -Arguments @(
    "-m", "pip", "install", "--disable-pip-version-check", "--upgrade", "pip==26.1.2"
)
Invoke-Native -FilePath $buildPython -Arguments @(
    "-m", "pip", "install", "--disable-pip-version-check", "--upgrade",
    "--only-binary=:all:", "-r", $requirementsFile
)

$artifactBaseName = (& $buildPython $versionFile).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($artifactBaseName)) {
    throw "Could not determine the application version from: $versionFile"
}
if ($artifactBaseName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
    throw "Version produced an invalid executable name: $artifactBaseName"
}
$artifactFileName = "$artifactBaseName.exe"
$stageExe = Join-Path $stageDist $artifactFileName
$publishedExe = Join-Path $distDirectory $artifactFileName

if (Test-Path -LiteralPath $stageDirectory) {
    Remove-Item -LiteralPath $stageDirectory -Recurse -Force
}
if (Test-Path -LiteralPath $workDirectory) {
    Remove-Item -LiteralPath $workDirectory -Recurse -Force
}
New-Item -ItemType Directory -Path $stageDirectory -Force | Out-Null

Write-Host "Checking GIS packages before packaging..." -ForegroundColor Cyan
try {
    Invoke-Native -FilePath $buildPython -Arguments @($selfTestScript, $pythonTestReport)
}
catch {
    Write-Host "GIS environment is damaged; reinstalling pinned packages once..." -ForegroundColor Yellow
    Invoke-Native -FilePath $buildPython -Arguments @(
        "-m", "pip", "install", "--disable-pip-version-check", "--force-reinstall",
        "--no-cache-dir", "--only-binary=:all:", "-r", $requirementsFile
    )
    Invoke-Native -FilePath $buildPython -Arguments @($selfTestScript, $pythonTestReport)
}

Write-Host "Building $artifactFileName..." -ForegroundColor Cyan
Push-Location $projectRoot
try {
    Invoke-Native -FilePath $buildPython -Arguments @(
        "-m", "PyInstaller", "--noconfirm", "--clean",
        "--distpath", $stageDist,
        "--workpath", $workDirectory,
        $specFile
    )
}
finally {
    Pop-Location
}

if (-not (Test-Path -LiteralPath $stageExe -PathType Leaf)) {
    throw "PyInstaller reported success, but the expected EXE is missing: $stageExe"
}

Write-Host "Testing native GIS I/O inside the finished EXE..." -ForegroundColor Cyan
$env:K_TOOLS_SELF_TEST_REPORT = $frozenTestReport
try {
    $process = Start-Process -FilePath $stageExe -ArgumentList "--build-self-test" -Wait -PassThru
}
finally {
    Remove-Item Env:K_TOOLS_SELF_TEST_REPORT -ErrorAction SilentlyContinue
}

if ($process.ExitCode -ne 0) {
    if (Test-Path -LiteralPath $frozenTestReport -PathType Leaf) {
        Get-Content -LiteralPath $frozenTestReport -Encoding UTF8 | Write-Host
    }
    throw "Frozen GIS self-test failed with exit code $($process.ExitCode). The previous dist was preserved."
}

# Publish only a build that passed the frozen MapInfo TAB round-trip test.
New-Item -ItemType Directory -Path $distDirectory -Force | Out-Null
Get-ChildItem -LiteralPath $distDirectory -Filter "K_Tools*.exe*" -File |
    Remove-Item -Force
Copy-Item -LiteralPath $stageExe -Destination $publishedExe -Force

$hash = Get-FileHash -LiteralPath $publishedExe -Algorithm SHA256
$hashLine = "$($hash.Hash.ToLowerInvariant())  $artifactFileName"
Set-Content -LiteralPath "$publishedExe.sha256" -Value $hashLine -Encoding ASCII

$sizeMb = [Math]::Round((Get-Item -LiteralPath $publishedExe).Length / 1MB, 1)
Write-Host ""
Write-Host "BUILD OK" -ForegroundColor Green
Write-Host "EXE: $publishedExe"
Write-Host "Size: $sizeMb MB"
Write-Host "GIS test: $frozenTestReport"
Write-Host "SHA256: $($hash.Hash.ToLowerInvariant())"
