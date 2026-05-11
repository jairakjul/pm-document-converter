param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$SpecPath = Join-Path $ProjectRoot "packaging\PMDocumentConverter.spec"

if (-not (Test-Path $PythonExe)) {
    throw "Python virtual environment not found: $PythonExe"
}

if (-not (Test-Path $SpecPath)) {
    throw "PyInstaller spec not found: $SpecPath"
}

Push-Location $ProjectRoot
try {
    if ($Clean) {
        $BuildDir = Join-Path $ProjectRoot "build"
        $DistDir = Join-Path $ProjectRoot "dist"

        if (Test-Path $BuildDir) {
            Remove-Item -LiteralPath $BuildDir -Recurse -Force
        }
        if (Test-Path $DistDir) {
            Remove-Item -LiteralPath $DistDir -Recurse -Force
        }
    }

    & $PythonExe -m PyInstaller --noconfirm --clean $SpecPath

    $ExePath = Join-Path $ProjectRoot "dist\PMDocumentConverter\PMDocumentConverter.exe"
    if (-not (Test-Path $ExePath)) {
        throw "Build finished but executable was not found: $ExePath"
    }

    $InternalDir = Join-Path $ProjectRoot "dist\PMDocumentConverter\_internal"
    $UnusedPaths = @(
        (Join-Path $InternalDir "Pythonwin"),
        (Join-Path $InternalDir "PIL\_avif.cp312-win_amd64.pyd"),
        (Join-Path $InternalDir "numpy-2.4.4.dist-info")
    )

    foreach ($Path in $UnusedPaths) {
        if (Test-Path $Path) {
            Remove-Item -LiteralPath $Path -Recurse -Force
        }
    }

    Write-Host "Build completed: $ExePath"
}
finally {
    Pop-Location
}
