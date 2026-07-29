param(
    [string]$NomeCompleto = "",
    [string]$Arquivo = "",
    [Alias("Diagnostic")]
    [switch]$Diagnostico,
    [switch]$Version,
    [switch]$SemPivotNativo,
    [switch]$Verbose,
    [Alias("h")]
    [switch]$Help
)

$ErrorActionPreference = "Stop"

function Show-Usage {
    Write-Host @"
Excel Compras Automation

Uso:
  .\run.ps1 -NomeCompleto "NOME SOBRENOME" [-Arquivo "entrada.xlsx"]
  .\run.ps1 -Diagnostico [-Arquivo "entrada.xlsx"]
  .\run.ps1 -Version
  .\run.ps1 -Help

Opcoes:
  -NomeCompleto       Nome usado no arquivo de saida. No modo normal, sera
                      solicitado se nao for informado.
  -Arquivo            Caminho opcional da planilha. Quando omitido, a aplicacao
                      usa a unica planilha existente na pasta input.
  -Diagnostico        Verifica o ambiente sem executar a automacao.
  -Diagnostic         Alias de -Diagnostico.
  -Version            Mostra a versao da aplicacao.
  -SemPivotNativo     Usa o resumo compativel em vez da PivotTable nativa.
  -Verbose            Mostra informacoes tecnicas adicionais.
  -Help, -h           Mostra esta ajuda.
"@
}

if ($Help) {
    Show-Usage
    exit 0
}

Set-Location -LiteralPath $PSScriptRoot
$PortableExecutable = Join-Path $PSScriptRoot "ExcelComprasAutomation.exe"
$PortableMarker = Join-Path $PSScriptRoot "VERSAO.txt"
$VenvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$UsePortableExecutable = Test-Path -LiteralPath $PortableExecutable -PathType Leaf

if (
    -not $UsePortableExecutable -and
    (Test-Path -LiteralPath $PortableMarker -PathType Leaf)
) {
    throw (
        "O pacote portatil esta incompleto: ExcelComprasAutomation.exe " +
        "nao foi encontrado. Extraia novamente o ZIP oficial inteiro."
    )
}

if (
    -not $UsePortableExecutable -and
    -not (Test-Path -LiteralPath $VenvPython -PathType Leaf)
) {
    throw (
        "Executavel portatil e ambiente virtual nao encontrados. Execute: " +
        "powershell -ExecutionPolicy Bypass -File .\setup.ps1"
    )
}

if ($Diagnostico -and $Version) {
    throw "Use apenas um modo por vez: -Diagnostico ou -Version."
}

$ApplicationArguments = @()

if (-not $UsePortableExecutable) {
    $ApplicationArguments += @("-m", "src.main")
}

if ($Version) {
    $ApplicationArguments += "--version"
}
else {
    if (-not [string]::IsNullOrWhiteSpace($Arquivo)) {
        $ApplicationArguments += @("--input", $Arquivo)
    }

    if ($Diagnostico) {
        $ApplicationArguments += "--diagnostico"
    }
    else {
        if ([string]::IsNullOrWhiteSpace($NomeCompleto)) {
            $NomeCompleto = Read-Host "Digite seu nome completo"
        }

        if ([string]::IsNullOrWhiteSpace($NomeCompleto)) {
            throw "Informe seu nome completo para executar a automacao."
        }

        $ApplicationArguments += @("--candidate-name", $NomeCompleto.Trim())
    }

    if ($SemPivotNativo) {
        $ApplicationArguments += "--sem-pivot-nativo"
    }

    if ($Verbose) {
        $ApplicationArguments += "--verbose"
    }
}

if ($UsePortableExecutable) {
    & $PortableExecutable @ApplicationArguments
}
else {
    & $VenvPython @ApplicationArguments
}
$ExitCode = $LASTEXITCODE

if ($null -eq $ExitCode) {
    throw "A aplicacao terminou sem informar um codigo de saida."
}

exit $ExitCode
