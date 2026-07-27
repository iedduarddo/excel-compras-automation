param(
    [string]$NomeCompleto = "",
    [string]$Arquivo = ".\input\teste_excel_analista_compras_celula_reservas.xlsx"
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

if (-not (Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    throw "Ambiente virtual não encontrado. Execute primeiro: powershell -ExecutionPolicy Bypass -File .\setup.ps1"
}

if ([string]::IsNullOrWhiteSpace($NomeCompleto)) {
    $NomeCompleto = Read-Host "Digite seu nome completo"
}

& ".\.venv\Scripts\python.exe" -m src.main `
    --input $Arquivo `
    --candidate-name $NomeCompleto

