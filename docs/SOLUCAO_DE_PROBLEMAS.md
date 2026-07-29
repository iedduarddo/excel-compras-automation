# Solução de problemas

## Pacote portátil

Se você baixou `ExcelComprasAutomation-vX.Y.Z-windows-x64.zip`, extraia todo o
conteúdo antes de executar. O pacote não precisa de Python, `.venv`,
`setup.ps1` ou internet.

Não mova somente `ExcelComprasAutomation.exe`. Ele precisa permanecer ao lado de
`_internal`, `config`, `run.ps1` e `iniciar.cmd`. Use uma pasta gravável do seu
usuário; evite executar dentro do ZIP ou de `Program Files`.

Confira primeiro:

```powershell
.\iniciar.cmd -Version
.\iniciar.cmd -Diagnostico
```

Se o Windows SmartScreen exibir um aviso, confirme que o ZIP veio da GitHub
Release oficial e compare seu SHA-256 com o arquivo `.sha256` publicado junto
dele. O executável ainda não possui assinatura de código.

As instruções de `.venv`, `setup.ps1` e instalação de bibliotecas abaixo se
aplicam apenas à execução pelo código-fonte.

## Primeira execução

O caminho recomendado é dar duplo clique em `iniciar.cmd`. Se a `.venv` ainda
não existir, o iniciador chama `setup.ps1`, que:

- cria `input`, `output`, `backup` e `logs`;
- localiza Python 3.11, 3.12, 3.13 ou 3.14;
- cria ou reutiliza uma `.venv` compatível;
- instala e verifica as dependências.

Quando chamado sem argumentos, `iniciar.cmd` mantém a janela aberta ao final para
que a mensagem possa ser lida. No terminal, argumentos são repassados para
`run.ps1` e não há pausa:

```powershell
.\iniciar.cmd -NomeCompleto "SEU NOME COMPLETO"
```

Consulte a ajuda sem preparar ou executar a automação:

```powershell
.\run.ps1 -Help
.\run.ps1 -h
```

## Diagnóstico básico

Execute, na raiz do projeto:

```powershell
.\run.ps1 -Diagnostico
```

O diagnóstico não pede nome e não cria backup, arquivo de saída ou log. Ele
verifica a versão do Python, as configurações, a entrada e as abas obrigatórias.
`[AVISO]` indica que o modo alternativo continua disponível; somente `[ERRO]`
faz o comando terminar com código 1.

Também é possível usar o alias em inglês:

```powershell
.\run.ps1 -Diagnostic
```

Consulte a versão do programa com:

```powershell
.\run.ps1 -Version
```

Esse comando também não cria backup, saída ou log. Se preferir passar pelo
iniciador, use `.\iniciar.cmd -Diagnostico` ou `.\iniciar.cmd -Version`; nesse
caso, uma `.venv` ausente será preparada antes.

Se o ambiente virtual ainda não existir, faça as verificações manuais:

```powershell
Get-Location
Get-ChildItem
py -3.14 --version
Test-Path .\.venv\Scripts\python.exe
Test-Path .\requirements.txt
Get-ChildItem .\input\*.xlsx, .\input\*.xlsm
```

Para conferir outra versão suportada, substitua `3.14` por `3.13`, `3.12` ou
`3.11`.

## Ambiente virtual

Se `.venv` não existe, execute:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

O setup rejeita versões anteriores ao Python 3.11 e posteriores ao 3.14. Ele
também interrompe imediatamente se a criação da `.venv`, o `pip` ou a verificação
das dependências terminar com erro.

Não é necessário ativar a `.venv` para usar `iniciar.cmd` ou `run.ps1`. Para
trabalhar manualmente no ambiente, a ativação opcional é:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\.venv\Scripts\Activate.ps1
```

Também é possível executar comandos Python sem ativar:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m src.main --nome "SEU NOME COMPLETO"
```

## Nenhuma planilha foi encontrada

Copie um arquivo `.xlsx` ou `.xlsm` para a pasta `input`. Arquivos temporários
que começam com `~$` são ignorados. Depois execute novamente `iniciar.cmd`.

## Há mais de uma planilha na pasta input

Escolha explicitamente a entrada:

```powershell
.\iniciar.cmd `
    -Arquivo ".\input\planilha_escolhida.xlsx" `
    -NomeCompleto "SEU NOME COMPLETO"
```

`-Arquivo` não é necessário quando existe uma única planilha válida.

## Bibliotecas

Confira:

```powershell
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -c "import openpyxl; print(openpyxl.__version__)"
.\.venv\Scripts\python.exe -c "import win32com.client; print('pywin32 OK')"
```

Se houver `ModuleNotFoundError`:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Não é necessário selecionar um interpretador no VS Code para executar o projeto.

## Arquivo em uso

Um `PermissionError` geralmente indica:

- entrada aberta no Excel;
- saída aberta no Excel;
- visualizador bloqueando o arquivo;
- antivírus analisando a pasta.

Feche o arquivo e execute novamente. O programa não sobrescreve uma saída
existente; ele adiciona horário ao nome.

## O Excel não liberou todos os recursos

A automação interrompe a execução de propósito para não reabrir uma saída que
ainda pode estar bloqueada. Aguarde alguns segundos para que o Windows conclua o
encerramento assíncrono e feche somente os arquivos ou as janelas do Excel
envolvidos nesta tentativa. Não encerre indiscriminadamente processos `EXCEL.EXE`
ou sessões preexistentes no Gerenciador de Tarefas. Depois, tente novamente.

Se o problema persistir, inicie uma nova execução no modo compatível:

```powershell
.\iniciar.cmd -SemPivotNativo
```

## Arquivo inválido

`BadZipFile` significa que o `.xlsx` está corrompido ou é outro formato
renomeado.

Baixe novamente o original. Não corrija somente trocando a extensão.

## Abas ou colunas ausentes

Abra o log mais recente:

```powershell
Get-ChildItem .\logs\*.log |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName
```

Se o nome mudou, adicione um alias. Se o dado realmente não existe, restaure uma
cópia completa.

## Excel Desktop e PivotTable

Teste a automação:

```powershell
.\.venv\Scripts\python.exe -c "import win32com.client; e=win32com.client.Dispatch('Excel.Application'); print(e.Version); e.Quit()"
```

Possíveis causas de falha:

- Excel Desktop não instalado;
- Office aguardando primeiro login;
- caixa de diálogo de licença;
- política corporativa bloqueando COM;
- Excel travado em segundo plano;
- arquivo em pasta de rede sem confiança.

Abra o Excel manualmente uma vez, conclua login/licença, feche e tente novamente.

Modo alternativo:

```powershell
.\iniciar.cmd -SemPivotNativo
```

## Fórmulas sem resultado imediato

O arquivo solicita recálculo automático. No Excel:

```text
Fórmulas → Opções de Cálculo → Automático
```

Depois:

```text
Ctrl + Alt + F9
```

Salve.

## Log detalhado

```powershell
.\iniciar.cmd -Verbose
```

Não publique logs sem revisar caminhos e nomes de arquivos.
