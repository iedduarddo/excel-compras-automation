# Solução de problemas

## Diagnóstico básico

Execute, na raiz:

```powershell
Get-Location
Get-ChildItem
py -3 --version
Test-Path .\.venv\Scripts\python.exe
Test-Path .\requirements.txt
Test-Path .\input\teste_excel_analista_compras_celula_reservas.xlsx
```

## Ambiente virtual

Se `.venv` não existe:

```powershell
py -3 -m venv .venv
```

Se a ativação é bloqueada:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\.venv\Scripts\Activate.ps1
```

Também é possível executar sem ativar:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m src.main --nome "SEU NOME COMPLETO"
```

## Bibliotecas

Confira:

```powershell
python -m pip check
python -c "import openpyxl; print(openpyxl.__version__)"
python -c "import win32com.client; print('pywin32 OK')"
```

Se houver `ModuleNotFoundError`:

```powershell
python -m pip install -r requirements.txt
```

Confirme antes que `sys.executable` aponta para `.venv`.

## Arquivo em uso

Um `PermissionError` geralmente indica:

- entrada aberta no Excel;
- saída aberta no Excel;
- visualizador bloqueando o arquivo;
- antivírus analisando a pasta.

Feche o arquivo e execute novamente. O programa não sobrescreve uma saída
existente; ele adiciona horário ao nome.

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
python -c "import win32com.client; e=win32com.client.Dispatch('Excel.Application'); print(e.Version); e.Quit()"
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
python -m src.main --nome "SEU NOME COMPLETO" --sem-pivot-nativo
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
python -m src.main --nome "SEU NOME COMPLETO" --verbose
```

Não publique logs sem revisar caminhos e nomes de arquivos.

