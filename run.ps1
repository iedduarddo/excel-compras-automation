param(
    [string]$NomeCompleto = "",
    [string]$Arquivo = ".\input\teste_excel_analista_compras_celula_reservas.xlsx",
    [Alias("Diagnostic")]
    [switch]$Diagnostico,
    [switch]$Version,
    [switch]$SemPivotNativo,
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

if (-not (Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    throw "Ambiente virtual não encontrado. Execute primeiro: powershell -ExecutionPolicy Bypass -File .\setup.ps1"
}

if ($Diagnostico -and $Version) {
    throw "Use apenas um modo por vez: -Diagnostico ou -Version."
}

if (-not $Diagnostico -and -not $Version -and [string]::IsNullOrWhiteSpace($NomeCompleto)) {
    $NomeCompleto = Read-Host "Digite seu nome completo"
}

$PythonArguments = @(
    "-m",
    "src.main"
)

if ($Version) {
    $PythonArguments += "--version"
}
else {
    $PythonArguments += @("--input", $Arquivo)

    if ($Diagnostico) {
        $PythonArguments += "--diagnostico"
    }
    else {
        $PythonArguments += @("--candidate-name", $NomeCompleto)
    }

    if ($SemPivotNativo) {
        $PythonArguments += "--sem-pivot-nativo"
    }

    if ($Verbose) {
        $PythonArguments += "--verbose"
    }
}

& ".\.venv\Scripts\python.exe" @PythonArguments
exit $LASTEXITCODE
