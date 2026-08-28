param(
    [ValidateSet(
        "Bundle",
        "AllWindows",
        "FullNoConsole",
        "FullConsole",
        "PartialNoConsole",
        "PartialConsole",
        "LinuxPartial"
    )]
    [string]$Target = "Bundle",

    [switch]$Clean,
    [switch]$IncludeLinux,
    [switch]$SkipPrinterCheck,
    [switch]$NoInstallPyInstaller,
    [string]$VenvPath = "B:\ambientes\fera"
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Dist = Join-Path $Root "dist"
$VenvPython = $null

Set-Location $Root

function Write-Step {
    param([string]$Message)
    Write-Host "[FERA build] $Message" -ForegroundColor Cyan
}

function Initialize-FeraEnvironment {
    $script:VenvPython = $null
    $candidateRoots = @(
        $VenvPath,
        "B:\ambientes\fera"
    ) | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -Unique

    foreach ($root in $candidateRoots) {
        $activateCandidates = @(
            (Join-Path $root "Scripts\Activate.ps1"),
            (Join-Path $root "script\activate.ps1"),
            (Join-Path $root "Scripts\activate.ps1")
        )
        foreach ($activate in $activateCandidates) {
            if (Test-Path -LiteralPath $activate) {
                Write-Step "Ativando ambiente FERA: $activate"
                . $activate
                break
            }
        }

        $pythonCandidates = @(
            (Join-Path $root "Scripts\python.exe"),
            (Join-Path $root "script\python.exe"),
            (Join-Path $root "python.exe")
        )
        foreach ($python in $pythonCandidates) {
            if (Test-Path -LiteralPath $python) {
                $script:VenvPython = $python
                Write-Step "Python do ambiente FERA: $script:VenvPython"
                return
            }
        }
    }

    throw "Ambiente FERA nao encontrado. Esperado em '$VenvPath' com Scripts\python.exe e Scripts\Activate.ps1."
}

function Ensure-PyInstaller {
    if ($NoInstallPyInstaller) {
        Write-Step "Pulando instalacao/verificacao do PyInstaller por -NoInstallPyInstaller"
        return
    }

    Write-Step "Verificando PyInstaller no ambiente FERA"
    & $VenvPython -m PyInstaller --version | Out-Null
    if ($LASTEXITCODE -eq 0) {
        return
    }

    Write-Step "Instalando PyInstaller no ambiente FERA"
    & $VenvPython -m pip install -U pyinstaller
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao instalar PyInstaller no ambiente FERA com exit code $LASTEXITCODE"
    }

    & $VenvPython -m PyInstaller --version | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller continua indisponivel no ambiente FERA apos instalacao."
    }
}

function Assert-7ZipRuntime {
    $runtime = Join-Path $Root "third_party\7zip"
    foreach ($name in @("7z.exe", "7z.dll", "License.txt")) {
        $path = Join-Path $runtime $name
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Runtime 7-Zip obrigatorio nao encontrado: $path"
        }
    }
    Write-Step "Runtime 7-Zip embarcado validado: $runtime"
}

function Assert-DistChild {
    param([string]$Path)
    $distFull = [System.IO.Path]::GetFullPath($Dist)
    $pathFull = [System.IO.Path]::GetFullPath($Path)
    if (-not $pathFull.StartsWith($distFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Path fora de dist bloqueado por seguranca: $Path"
    }
}

function Remove-DistFolder {
    param([string]$Name)
    $path = Join-Path $Dist $Name
    Assert-DistChild $path
    if (Test-Path -LiteralPath $path) {
        Write-Step "Removendo $path"
        Remove-Item -LiteralPath $path -Recurse -Force
    }
}

function Invoke-PyInstallerSpec {
    param([string]$Spec)
    if (-not (Test-Path -LiteralPath (Join-Path $Root $Spec))) {
        throw "Spec nao encontrado: $Spec"
    }
    $arguments = @("--noconfirm")
    if ($Clean) {
        $arguments += "--clean"
    }
    $arguments += $Spec
    Write-Step "$VenvPython -m PyInstaller $($arguments -join ' ')"
    & $VenvPython -m PyInstaller @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller falhou para $Spec com exit code $LASTEXITCODE"
    }
}

function Copy-Tree {
    param(
        [string]$Source,
        [string]$Destination,
        [bool]$Required = $true
    )
    if (-not (Test-Path -LiteralPath $Source)) {
        if ($Required) {
            throw "Diretorio obrigatorio nao encontrado: $Source"
        }
        Write-Host "[FERA build] Aviso: diretorio opcional nao encontrado: $Source" -ForegroundColor Yellow
        return
    }
    if (Test-Path -LiteralPath $Destination) {
        Remove-Item -LiteralPath $Destination -Recurse -Force
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $Destination) -Force | Out-Null
    Copy-Item -LiteralPath $Source -Destination $Destination -Recurse -Force
}

function Copy-FileIfExists {
    param(
        [string]$Source,
        [string]$Destination,
        [bool]$Required = $false
    )
    if (-not (Test-Path -LiteralPath $Source)) {
        if ($Required) {
            throw "Arquivo obrigatorio nao encontrado: $Source"
        }
        Write-Host "[FERA build] Aviso: arquivo opcional nao encontrado: $Source" -ForegroundColor Yellow
        return
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $Destination) -Force | Out-Null
    Copy-Item -LiteralPath $Source -Destination $Destination -Force
}

function Add-SharedPayload {
    param(
        [string]$TargetName,
        [string]$PrinterFolderName
    )
    $targetPath = Join-Path $Dist $TargetName
    if (-not (Test-Path -LiteralPath $targetPath)) {
        throw "Destino nao encontrado para payload compartilhado: $targetPath"
    }

    $printerSource = Join-Path $Dist $PrinterFolderName
    $printerDestination = Join-Path $targetPath $PrinterFolderName
    Copy-Tree $printerSource $printerDestination (-not $SkipPrinterCheck)

    Copy-FileIfExists (Join-Path $Root "FERA.pdf") (Join-Path $targetPath "FERA.pdf") $true

    $callers = Join-Path $targetPath "FERA_CALLERS"
    New-Item -ItemType Directory -Path $callers -Force | Out-Null
    Copy-FileIfExists (Join-Path $Root "FERA-Linux.sh") (Join-Path $callers "FERA-Linux.sh")
    Copy-FileIfExists (Join-Path $Root "FERA-Windows.exe") (Join-Path $callers "FERA-Windows.exe")
    Copy-FileIfExists (Join-Path $Root "FERA.pdf") (Join-Path $callers "FERA.pdf") $true
}

function Publish-PyInstallerOutput {
    param(
        [string]$SourceName,
        [string]$TargetName,
        [string]$PrinterFolderName
    )
    $sourcePath = Join-Path $Dist $SourceName
    $targetPath = Join-Path $Dist $TargetName
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Saida esperada do PyInstaller nao encontrada: $sourcePath"
    }
    Remove-DistFolder $TargetName
    Write-Step "Movendo $sourcePath para $targetPath"
    Move-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    Add-SharedPayload $TargetName $PrinterFolderName
}

function Build-PartialWindows {
    param([bool]$Console)
    $spec = if ($Console) { "fera_partial_console.spec" } else { "fera_partial_noconsole.spec" }
    Invoke-PyInstallerSpec $spec
    Publish-PyInstallerOutput "fera" "FERA_WINDOWS" "printer_interface_windows"
}

function Build-FullWindows {
    param([bool]$Console)
    $spec = if ($Console) { "fera_console.spec" } else { "fera_noconsole.spec" }
    Invoke-PyInstallerSpec $spec
    Publish-PyInstallerOutput "fera_full" "FERA_FULL_WINDOWS" "printer_interface_windows"
}

function Build-LinuxPartial {
    Invoke-PyInstallerSpec "feralinux_partial_noconsole.spec"
    Publish-PyInstallerOutput "fera" "FERA_LINUX" "printer_interface_linux"
}

function Update-FullBundle {
    $full = Join-Path $Dist "FERA_FULL_WINDOWS"
    if (-not (Test-Path -LiteralPath $full)) {
        Write-Host "[FERA build] Aviso: FERA_FULL_WINDOWS nao existe; bundle nao atualizado." -ForegroundColor Yellow
        return
    }

    $windows = Join-Path $Dist "FERA_WINDOWS"
    if (Test-Path -LiteralPath $windows) {
        Copy-Tree $windows (Join-Path $full "FERA_WINDOWS") $true
    } else {
        Write-Host "[FERA build] Aviso: FERA_WINDOWS nao existe para embutir no bundle full." -ForegroundColor Yellow
    }

    $linux = Join-Path $Dist "FERA_LINUX"
    if (Test-Path -LiteralPath $linux) {
        Copy-Tree $linux (Join-Path $full "FERA_LINUX") $true
    } else {
        Write-Host "[FERA build] Aviso: FERA_LINUX nao existe para embutir no bundle full." -ForegroundColor Yellow
    }

    Add-SharedPayload "FERA_FULL_WINDOWS" "printer_interface_windows"
}

New-Item -ItemType Directory -Path $Dist -Force | Out-Null
Initialize-FeraEnvironment
Ensure-PyInstaller
Assert-7ZipRuntime

switch ($Target) {
    "PartialNoConsole" {
        Build-PartialWindows $false
    }
    "PartialConsole" {
        Build-PartialWindows $true
    }
    "FullNoConsole" {
        Build-FullWindows $false
        Update-FullBundle
    }
    "FullConsole" {
        Build-FullWindows $true
        Update-FullBundle
    }
    "LinuxPartial" {
        Build-LinuxPartial
    }
    "AllWindows" {
        Build-PartialWindows $false
        Build-FullWindows $false
        Update-FullBundle
    }
    "Bundle" {
        Build-PartialWindows $false
        if ($IncludeLinux) {
            Build-LinuxPartial
        }
        Build-FullWindows $false
        Update-FullBundle
    }
}

Write-Step "Concluido. Saidas em $Dist"
