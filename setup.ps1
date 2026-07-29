param()

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$VenvDirectory = Join-Path $PSScriptRoot ".venv"
$VenvPython = Join-Path $VenvDirectory "Scripts\python.exe"
$RequirementsFile = Join-Path $PSScriptRoot "requirements.txt"
$SupportedPythonProbe = "import sys; version = sys.version_info; print('.'.join(map(str, version[:3]))); raise SystemExit(0 if (3, 11) <= version[:2] < (3, 15) else 3)"

function Invoke-ExternalCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$ArgumentList = @(),

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    & $FilePath @ArgumentList
    $ExitCode = $LASTEXITCODE

    if ($ExitCode -ne 0) {
        throw "$Description falhou com o codigo de saida $ExitCode."
    }
}

function Get-SupportedPythonVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$Prefix = @()
    )

    $ProbeArguments = @($Prefix) + @("-I", "-c", $SupportedPythonProbe)
    try {
        $ProbeOutput = & $FilePath @ProbeArguments 2>$null
        $ProbeExitCode = $LASTEXITCODE
    }
    catch {
        return $null
    }

    if ($ProbeExitCode -ne 0) {
        return $null
    }

    return [string]($ProbeOutput | Select-Object -Last 1)
}

function Find-SupportedPython {
    $Candidates = @()
    $PyLauncher = Get-Command -Name "py.exe" -CommandType Application `
        -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($null -ne $PyLauncher) {
        foreach ($Version in @("3.14", "3.13", "3.12", "3.11")) {
            $Candidates += [pscustomobject]@{
                FilePath = $PyLauncher.Source
                Prefix   = @("-$Version")
                Label    = "py -$Version"
            }
        }
    }

    $PythonCommand = Get-Command -Name "python.exe" -CommandType Application `
        -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($null -ne $PythonCommand) {
        $Candidates += [pscustomobject]@{
            FilePath = $PythonCommand.Source
            Prefix   = @()
            Label    = "python"
        }
    }

    foreach ($Candidate in $Candidates) {
        $DetectedVersion = Get-SupportedPythonVersion `
            -FilePath $Candidate.FilePath `
            -Prefix $Candidate.Prefix

        if ($null -ne $DetectedVersion) {
            return [pscustomobject]@{
                FilePath = $Candidate.FilePath
                Prefix   = $Candidate.Prefix
                Label    = $Candidate.Label
                Version  = $DetectedVersion
            }
        }
    }

    throw (
        "Python compativel nao foi encontrado. Instale o Python >=3.11 e <3.15 " +
        "e reabra o terminal."
    )
}

Write-Host ""
Write-Host "Excel Compras Automation - preparacao do ambiente" -ForegroundColor Cyan
Write-Host ""

Write-Host "1/5 Criando as pastas operacionais..."
foreach ($DirectoryName in @("input", "output", "backup", "logs")) {
    $DirectoryPath = Join-Path $PSScriptRoot $DirectoryName
    if (-not (Test-Path -LiteralPath $DirectoryPath -PathType Container)) {
        New-Item -ItemType Directory -Path $DirectoryPath -Force | Out-Null
    }
}

if (-not (Test-Path -LiteralPath $RequirementsFile -PathType Leaf)) {
    throw "Arquivo de dependencias nao encontrado: $RequirementsFile"
}

if (Test-Path -LiteralPath $VenvPython -PathType Leaf) {
    Write-Host "2/5 Verificando o Python da .venv existente..."
    $VenvVersion = Get-SupportedPythonVersion -FilePath $VenvPython

    if ($null -eq $VenvVersion) {
        throw (
            "A .venv existente nao usa um Python suportado ou esta corrompida. " +
            "Ela foi preservada. Use Python >=3.11 e <3.15 para repara-la."
        )
    }

    Write-Host "Python $VenvVersion encontrado na .venv."
    Write-Host "3/5 A .venv ja existe e sera reutilizada."
}
else {
    Write-Host "2/5 Localizando um Python compativel..."
    $BootstrapPython = Find-SupportedPython
    Write-Host "Python $($BootstrapPython.Version) encontrado via $($BootstrapPython.Label)."

    if (Test-Path -LiteralPath $VenvDirectory -PathType Container) {
        Write-Host "3/5 Completando a .venv existente sem remove-la..."
    }
    else {
        Write-Host "3/5 Criando a .venv..."
    }

    $VenvArguments = @($BootstrapPython.Prefix) + @("-m", "venv", $VenvDirectory)
    Invoke-ExternalCommand `
        -FilePath $BootstrapPython.FilePath `
        -ArgumentList $VenvArguments `
        -Description "Criacao do ambiente virtual"

    if (-not (Test-Path -LiteralPath $VenvPython -PathType Leaf)) {
        throw "A criacao da .venv terminou sem produzir: $VenvPython"
    }

    $VenvVersion = Get-SupportedPythonVersion -FilePath $VenvPython
    if ($null -eq $VenvVersion) {
        throw "A .venv foi criada, mas o Python dela nao esta operacional."
    }
}

Write-Host "4/5 Atualizando o pip..."
Invoke-ExternalCommand `
    -FilePath $VenvPython `
    -ArgumentList @("-m", "pip", "install", "--upgrade", "pip") `
    -Description "Atualizacao do pip"

Write-Host "5/5 Instalando e verificando as dependencias..."
Invoke-ExternalCommand `
    -FilePath $VenvPython `
    -ArgumentList @("-m", "pip", "install", "-r", $RequirementsFile) `
    -Description "Instalacao das dependencias"
Invoke-ExternalCommand `
    -FilePath $VenvPython `
    -ArgumentList @("-m", "pip", "check") `
    -Description "Verificacao das dependencias"

Write-Host ""
Write-Host "Ambiente preparado com sucesso." -ForegroundColor Green
Write-Host "Python da aplicacao: $VenvPython"
