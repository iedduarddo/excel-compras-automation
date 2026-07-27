param()

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

Write-Host ""
Write-Host "Excel Compras Automation - preparação do ambiente" -ForegroundColor Cyan
Write-Host ""

if (Get-Command py -ErrorAction SilentlyContinue) {
    $PythonCommand = "py"
    $PythonPrefix = @("-3")
}
elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $PythonCommand = "python"
    $PythonPrefix = @()
}
else {
    throw "Python não foi encontrado. Instale o Python 3 e reabra o terminal."
}

Write-Host "1/4 Verificando o Python..."
& $PythonCommand @PythonPrefix --version

if (-not (Test-Path -LiteralPath ".\.venv\Scripts\python.exe")) {
    Write-Host "2/4 Criando o ambiente virtual .venv..."
    & $PythonCommand @PythonPrefix -m venv .venv
}
else {
    Write-Host "2/4 O ambiente virtual .venv já existe."
}

Write-Host "3/4 Atualizando o pip..."
& ".\.venv\Scripts\python.exe" -m pip install --upgrade pip

Write-Host "4/4 Instalando as bibliotecas do projeto..."
& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt
& ".\.venv\Scripts\python.exe" -m pip check

Write-Host ""
Write-Host "Ambiente preparado com sucesso." -ForegroundColor Green
Write-Host "No VS Code, selecione: .venv\Scripts\python.exe"
