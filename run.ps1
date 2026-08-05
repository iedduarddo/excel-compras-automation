param(
    [string]$NomeCompleto = "",
    [string]$Arquivo = "",
    [string]$Adaptador = "",
    [switch]$Lote,
    [switch]$Assistente,
    [string]$Comando = "",
    [string]$PastaAssistente = "",
    [switch]$Monitorar,
    [switch]$PrepararPastas,
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
  .\run.ps1 -Lote -NomeCompleto "NOME SOBRENOME"
  .\run.ps1 -Assistente -PrepararPastas
  .\run.ps1 -Assistente -Comando 'diagnosticar todas'
  .\run.ps1 -Assistente -Monitorar
  .\run.ps1 -Diagnostico [-Arquivo "entrada.xlsx"] [-Adaptador "perfil.json"]
  .\run.ps1 -Version
  .\run.ps1 -Help

Opcoes:
  -NomeCompleto       Nome usado no arquivo de saida. No modo normal, sera
                      solicitado se nao for informado.
  -Arquivo            Caminho opcional da planilha. Quando omitido, a aplicacao
                      usa a unica planilha existente na pasta input.
  -Lote               Processa todas as planilhas validas da pasta input.
  -Adaptador          Perfil JSON opcional com aliases especificos da origem.
  -Assistente         Usa a pasta monitorada e a fila de comandos escritos.
  -Comando            Executa um comando conhecido diretamente.
  -PastaAssistente    Define uma raiz alternativa para as pastas monitoradas.
  -Monitorar          Mantem o assistente aguardando entradas e comandos.
  -PrepararPastas     Cria a estrutura do assistente sem processar arquivos.
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

if ($Lote -and ($Diagnostico -or $Version)) {
    throw "Use apenas um modo por vez: -Lote, -Diagnostico ou -Version."
}

if ($Lote -and -not [string]::IsNullOrWhiteSpace($Arquivo)) {
    throw "Nao combine -Lote com -Arquivo. O lote usa a pasta input."
}

if (
    (
        -not [string]::IsNullOrWhiteSpace($Comando) -or
        -not [string]::IsNullOrWhiteSpace($PastaAssistente) -or
        $Monitorar -or $PrepararPastas
    ) -and
    -not $Assistente
) {
    throw "Use -Assistente junto de -Comando, -Monitorar ou -PrepararPastas."
}

$QuantidadeAcoesAssistente = 0

if (-not [string]::IsNullOrWhiteSpace($Comando)) {
    $QuantidadeAcoesAssistente++
}

if ($Monitorar) {
    $QuantidadeAcoesAssistente++
}

if ($PrepararPastas) {
    $QuantidadeAcoesAssistente++
}

if ($QuantidadeAcoesAssistente -gt 1) {
    throw "Use apenas uma acao do assistente por execucao."
}

if (
    $Assistente -and
    (
        $Lote -or $Diagnostico -or
        -not [string]::IsNullOrWhiteSpace($Arquivo) -or
        -not [string]::IsNullOrWhiteSpace($NomeCompleto) -or
        -not [string]::IsNullOrWhiteSpace($Adaptador)
    )
) {
    throw "O assistente usa sua propria pasta e nao aceita os modos tradicionais."
}

$ApplicationArguments = @()

if (-not $UsePortableExecutable) {
    $ApplicationArguments += @("-m", "src.main")
}

if ($Version) {
    $ApplicationArguments += "--version"
}
else {
    if ($Assistente) {
        $ApplicationArguments += "--assistente"

        if (-not [string]::IsNullOrWhiteSpace($PastaAssistente)) {
            $ApplicationArguments += @("--pasta-assistente", $PastaAssistente)
        }

        if (-not [string]::IsNullOrWhiteSpace($Comando)) {
            $ApplicationArguments += @("--comando", $Comando)
        }

        if ($Monitorar) {
            $ApplicationArguments += "--monitorar"
        }

        if ($PrepararPastas) {
            $ApplicationArguments += "--preparar-pastas"
        }
    }
    else {
        if (-not [string]::IsNullOrWhiteSpace($Arquivo)) {
            $ApplicationArguments += @("--input", $Arquivo)
        }

        if (-not [string]::IsNullOrWhiteSpace($Adaptador)) {
            $ApplicationArguments += @("--adaptador", $Adaptador)
        }

        if ($Diagnostico) {
            $ApplicationArguments += "--diagnostico"
        }
        else {
            if ($Lote) {
                $ApplicationArguments += "--lote"
            }

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
