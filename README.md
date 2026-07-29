# Excel Compras Automation

[![CI](https://github.com/iedduarddo/excel-compras-automation/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/iedduarddo/excel-compras-automation/actions/workflows/ci.yml)

Automação em Python para o **Teste Prático de Excel — Analista de Compras |
Célula de Reservas**.

Este projeto foi preparado para Windows 11 e para uma pessoa que nunca utilizou
Python. O Visual Studio Code é opcional: ele ajuda a estudar ou editar o projeto,
mas não é necessário para executar a automação. Não é preciso entender todo o
código antes da primeira execução.

O programa:

- preserva a planilha original;
- cria um backup antes de processar;
- identifica abas e colunas mesmo quando nomes e posições mudam;
- preenche as colunas calculadas com fórmulas do Excel;
- calcula os dez indicadores da aba `Respostas`;
- identifica e justifica as cinco solicitações mais urgentes;
- aplica formatação condicional;
- cria uma Tabela Dinâmica nativa e um gráfico quando o Excel Desktop está
  disponível;
- cria um resumo formula-driven como alternativa de compatibilidade;
- grava logs e valida o arquivo antes de informar sucesso.

## Resultado que será criado

O arquivo final será salvo na pasta `output` com o nome completo informado:

```text
Seu_Nome_Completo_Teste_Excel_Analista_Compras.xlsx
```

O arquivo da pasta `input` não é alterado.

---

# Início rápido

Instale uma versão do Python entre 3.11 e 3.14, extraia o projeto do `.zip` e
coloque uma única planilha `.xlsx` ou `.xlsm` na pasta `input`. Depois, na pasta
principal, dê duplo clique em:

```text
iniciar.cmd
```

Na primeira execução, o iniciador cria as pastas operacionais, localiza um Python
compatível, cria a `.venv` e instala as dependências. Nas próximas execuções, a
mesma `.venv` é verificada e reutilizada. Quando aberto por duplo clique, o
iniciador mantém a janela aberta ao final para que as mensagens possam ser lidas.

O programa encontra automaticamente a única planilha e solicita o nome completo.
Se houver mais de uma, ele interrompe a execução, lista as entradas e orienta
como indicar a escolhida pelo terminal:

```powershell
.\iniciar.cmd -Arquivo ".\input\minha_planilha.xlsx"
```

Com argumentos, o iniciador não pausa ao final e devolve o código de saída ao
terminal. Consulte todas as opções com:

```powershell
.\run.ps1 -Help
.\run.ps1 -h
```

O diagnóstico verifica o ambiente e a estrutura da planilha sem pedir nome e sem
criar backup, saída ou log. A consulta de versão também não cria artefatos
operacionais:

```powershell
.\iniciar.cmd -Diagnostico
.\iniciar.cmd -Version
```

`-Diagnostic` é um alias equivalente a `-Diagnostico`. Se a `.venv` ainda não
existir, `iniciar.cmd` prepara o ambiente antes de encaminhar qualquer opção.

---

# 1. O que você precisa antes de começar

Você precisará de:

- Windows 11;
- acesso à internet durante a instalação;
- Python 3.11, 3.12, 3.13 ou 3.14;
- Microsoft Excel Desktop para criar a Tabela Dinâmica nativa;
- a planilha do teste em formato `.xlsx` ou `.xlsm`;
- alguns minutos para instalar as bibliotecas.

O Visual Studio Code e suas extensões são opcionais e recomendados somente para
quem deseja estudar, demonstrar ou modificar o código.

O Excel não precisa ficar aberto. A automação abre uma instância invisível,
cria a Tabela Dinâmica a partir de uma fonte estática validada, configura as
fórmulas para recálculo automático na próxima abertura e solicita o encerramento
do programa. O Windows pode concluir a liberação do processo de forma assíncrona.

Importante:

- não execute o projeto dentro do arquivo `.zip`;
- não abra a planilha de saída enquanto a automação estiver sendo executada;
- execute os comandos na pasta que contém este `README.md`;
- copie apenas o conteúdo dentro das caixas de código;
- não copie o texto `PS C:\...>` que aparece antes do cursor.

---

# 2. Instalar o Python no Windows 11

## 2.1. Abrir o site oficial

Abra:

https://www.python.org/downloads/windows/

Escolha uma versão estável do Python 3 para Windows. Este projeto aceita Python
3.11, 3.12, 3.13 e 3.14.

Para a maioria dos computadores, escolha:

```text
Windows installer (64-bit)
```

Não escolha:

- Alpha;
- Beta;
- Release Candidate ou RC;
- Embeddable package;
- pacote de código-fonte;
- versão de 32 bits em um computador comum de 64 bits.

Para descobrir a arquitetura do computador:

1. abra `Configurações`;
2. clique em `Sistema`;
3. clique em `Sobre`;
4. localize `Tipo de sistema`.

Se aparecer `x64`, use o instalador de 64 bits. Se aparecer `ARM64`, use o
instalador ARM64.

## 2.2. Executar o instalador

Na primeira tela do instalador tradicional:

1. marque `Add python.exe to PATH`;
2. mantenha o Python Launcher habilitado, se a opção aparecer;
3. clique em `Install Now`;
4. aguarde o término;
5. clique em `Disable path length limit`, se a opção aparecer;
6. feche o instalador.

Por que marcar `Add python.exe to PATH`?

Essa opção permite que o Windows encontre o Python quando você digitar um
comando no terminal.

## 2.3. Fechar terminais antigos

Depois da instalação:

1. feche o Prompt de Comando;
2. feche o PowerShell;
3. feche o VS Code, se ele já estiver aberto;
4. abra um novo PowerShell.

Um terminal que já estava aberto pode não reconhecer a instalação.

## 2.4. Verificar o Python

Execute:

```powershell
py -3.14 --version
```

O resultado deverá ser parecido com:

```text
Python 3.13.14
```

Você também pode substituir `3.14` por `3.13`, `3.12` ou `3.11`. O número de
correção pode ser diferente.

Verifique o instalador de bibliotecas:

```powershell
py -3 -m pip --version
```

O resultado deverá mencionar `pip` e um caminho do Python.

Veja todas as instalações localizadas:

```powershell
py -0p
```

Se esses três comandos funcionarem, o Python está pronto.

## 2.5. Se `python` abrir a Microsoft Store

Prefira o comando `py -3`. Se também quiser corrigir os aliases:

1. abra `Configurações`;
2. acesse `Aplicativos`;
3. acesse `Configurações avançadas de aplicativos`;
4. abra `Aliases de execução de aplicativo`;
5. desative os aliases `python.exe` e `python3.exe` da Microsoft Store;
6. feche e abra o terminal.

---

# 3. Opcional: instalar o Visual Studio Code

O VS Code não é necessário para usar `iniciar.cmd`. Instale-o se quiser estudar
o código, executar testes ou apresentar os detalhes técnicos do projeto.

## 3.1. Baixar

Abra o site oficial:

https://code.visualstudio.com/download

Para a maioria dos computadores Windows, baixe:

```text
User Installer x64
```

O `User Installer` é adequado para uma conta individual e normalmente não exige
permissão de administrador.

## 3.2. Instalar

Durante a instalação, marque, quando essas opções aparecerem:

- `Add to PATH`;
- `Register Code as an editor for supported file types`;
- `Add "Open with Code" action`;
- criar um atalho na Área de Trabalho, se desejar.

Conclua a instalação e abra o VS Code.

Não é necessário executar o VS Code como administrador.

## 3.3. Verificar pelo terminal

Feche e abra um novo PowerShell. Execute:

```powershell
code --version
```

Se aparecer uma versão, o comando está pronto.

Se `code` não for reconhecido, ainda será possível usar o VS Code normalmente
pelo menu Iniciar.

---

# 4. Opcional: instalar as extensões do VS Code

As extensões adicionam suporte a Python dentro do editor, mas não são usadas pelo
iniciador da automação.

## 4.1. Abrir a tela de extensões

No VS Code:

1. clique no ícone `Extensions` na barra lateral esquerda; ou
2. pressione `Ctrl + Shift + X`.

## 4.2. Extensões recomendadas para desenvolvimento

Pesquise e instale:

```text
Python
Publicador: Microsoft
ID: ms-python.python
```

Depois:

```text
Pylance
Publicador: Microsoft
ID: ms-python.vscode-pylance
```

O projeto já contém `.vscode/extensions.json`. Quando a pasta for aberta, o VS
Code também poderá sugerir essas extensões.

## 4.3. Extensões opcionais

Você pode instalar Jupyter e um visualizador de Excel, mas eles não são
necessários para executar este programa.

Pelo terminal, as extensões recomendadas também podem ser instaladas com:

```powershell
code --install-extension ms-python.python
code --install-extension ms-python.vscode-pylance
```

---

# 5. Extrair e abrir o projeto

## 5.1. Extrair o arquivo

Se o projeto foi entregue como `.zip`:

1. clique com o botão direito no arquivo;
2. escolha `Extrair Tudo`;
3. escolha uma pasta local;
4. confirme a extração.

Uma localização simples é:

```text
C:\Users\SEU_USUARIO\Projects\ExcelComprasAutomation
```

Não execute o projeto diretamente de dentro do `.zip`.

## 5.2. Opcional: abrir a pasta no VS Code

Se estiver usando o VS Code:

1. clique em `File`;
2. clique em `Open Folder`;
3. selecione a pasta `ExcelComprasAutomation`;
4. clique em `Select Folder`;
5. confirme que `README.md` e `requirements.txt` aparecem no Explorer.

Se o VS Code perguntar se você confia nos arquivos, escolha confiar porque esta
é a pasta do seu próprio projeto.

## 5.3. Opcional: abrir o terminal integrado

No VS Code:

```text
Terminal → New Terminal
```

Confirme a localização:

```powershell
Get-Location
```

Confira os arquivos:

```powershell
Get-ChildItem
```

Verifique o arquivo de dependências:

```powershell
Test-Path .\requirements.txt
```

Resultado esperado:

```text
True
```

Se aparecer `False`, a pasta errada foi aberta.

---

# 6. Preparar o ambiente automaticamente

O método recomendado é dar duplo clique em `iniciar.cmd`. Quando a `.venv` não
existe, ele chama `setup.ps1` automaticamente. Também é possível executar apenas
a preparação pelo terminal:

Na raiz do projeto:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

O script:

1. cria, se necessário, `input`, `output`, `backup` e `logs`;
2. localiza uma versão suportada do Python, entre 3.11 e 3.14;
3. cria ou valida a pasta `.venv`, sem apagar um ambiente existente;
4. atualiza o `pip`;
5. instala as bibliotecas de `requirements.txt` e executa `pip check`.

Cada comando externo tem seu código de saída verificado. Se uma etapa falhar, o
script interrompe a preparação e não informa um falso sucesso. Executá-lo
novamente é seguro: as pastas e uma `.venv` compatível são reutilizadas.

O final esperado é:

```text
Ambiente preparado com sucesso.
```

## 6.1. O que é `.venv`

`.venv` é um ambiente virtual. Ele mantém as bibliotecas deste projeto separadas
das bibliotecas de outros projetos.

Não envie a pasta `.venv` para o GitHub. Ela pode ser recriada com
`requirements.txt`.

---

# 7. Preparar o ambiente manualmente

Use este capítulo se quiser entender cada comando ou se `setup.ps1` não puder
ser utilizado. A ativação manual da `.venv` não é necessária para `iniciar.cmd`
ou `run.ps1`.

## 7.1. Criar o ambiente

```powershell
py -3.14 -m venv .venv
```

Se o Python 3.14 não estiver instalado, use `-3.13`, `-3.12` ou `-3.11`.

É normal o comando terminar sem mostrar mensagem.

## 7.2. Ativar

Esta etapa é opcional e serve somente para executar comandos Python manualmente:

```powershell
.\.venv\Scripts\Activate.ps1
```

O início do terminal deverá ficar parecido com:

```text
(.venv) PS C:\Users\...\ExcelComprasAutomation>
```

## 7.3. Corrigir o bloqueio de scripts

Se aparecer `running scripts is disabled`, execute:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
```

Depois:

```powershell
.\.venv\Scripts\Activate.ps1
```

Essa alteração vale somente para o terminal atual. Ela desaparece quando o
terminal é fechado.

Não use uma política `Unrestricted` para o computador inteiro.

## 7.4. Verificar o Python do ambiente

```powershell
.\.venv\Scripts\python.exe -c "import sys; print(sys.executable)"
```

O caminho deverá terminar em:

```text
ExcelComprasAutomation\.venv\Scripts\python.exe
```

Também é possível verificar:

```powershell
Get-Command .\.venv\Scripts\python.exe | Select-Object -ExpandProperty Source
```

O caminho deverá apontar para `.venv`.

## 7.5. Instalar as bibliotecas

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pip
```

Depois:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Confira conflitos:

```powershell
.\.venv\Scripts\python.exe -m pip check
```

Resultado ideal:

```text
No broken requirements found.
```

Teste as duas bibliotecas principais:

```powershell
.\.venv\Scripts\python.exe -c "import openpyxl, win32com.client; print('Bibliotecas instaladas com sucesso.')"
```

## 7.6. Por que há poucas bibliotecas

O projeto instala somente o que utiliza:

- `openpyxl`: leitura, fórmulas, formatação e gráfico de compatibilidade;
- `pywin32`: automação do Excel Desktop e Tabela Dinâmica nativa.

Pandas, NumPy, Matplotlib e XlsxWriter não são necessários nesta versão.
Instalar pacotes sem uso aumenta o tempo de instalação e a superfície de
manutenção.

Para executar os testes, instale:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

---

# 8. Opcional: selecionar o interpretador no VS Code

Esta configuração é útil para desenvolvimento. Ela não é necessária para
executar `iniciar.cmd`, `setup.ps1` ou `run.ps1`.

Criar o ambiente e selecionar o interpretador são etapas diferentes.

No VS Code:

1. pressione `Ctrl + Shift + P`;
2. digite `Python: Select Interpreter`;
3. abra o comando;
4. selecione a opção semelhante a:

```text
Python 3.x.x ('.venv': venv)
```

O caminho deverá terminar em:

```text
.venv\Scripts\python.exe
```

Se a opção não aparecer:

1. escolha `Enter interpreter path`;
2. clique em `Find`;
3. abra `.venv`;
4. abra `Scripts`;
5. selecione `python.exe`.

Depois, feche o terminal atual e abra outro.

---

# 9. Preparar a planilha de entrada

Coloque na pasta `input` a planilha que será processada. São aceitos arquivos
`.xlsx` e `.xlsm`, com qualquer nome. Exemplo:

```text
input\minha_planilha_de_compras.xlsx
```

Quando existe uma única planilha válida, ela é selecionada automaticamente.

Para trabalhar com outra cópia:

1. feche a planilha no Excel;
2. copie o arquivo original;
3. cole a cópia dentro de `input`;
4. mantenha a extensão `.xlsx` ou `.xlsm`.

Pelo PowerShell:

```powershell
Copy-Item -LiteralPath "$env:USERPROFILE\Downloads\minha_planilha.xlsx" `
    -Destination ".\input\"
```

Confirme:

```powershell
Get-ChildItem ".\input\*.xlsx", ".\input\*.xlsm" |
    Select-Object Name, Length, LastWriteTime
```

Não selecione arquivos temporários que começam com `~$`.
Se houver mais de uma planilha válida, use `-Arquivo` para indicar explicitamente
qual delas será processada.

---

# 10. Executar o projeto

## 10.1. Forma mais fácil

Com uma única planilha na pasta `input`, dê duplo clique em `iniciar.cmd` ou
execute:

```powershell
.\iniciar.cmd
```

O iniciador prepara o ambiente quando necessário, executa `run.ps1` e solicita o
nome completo. Para informar os valores diretamente:

```powershell
.\iniciar.cmd -NomeCompleto "Maria Aparecida da Silva"
```

Use `-Arquivo` somente quando houver mais de uma planilha, quando a entrada
estiver fora de `input` ou quando quiser escolher explicitamente:

```powershell
.\iniciar.cmd `
    -NomeCompleto "Maria Aparecida da Silva" `
    -Arquivo ".\input\compras_julho.xlsm"
```

Use aspas quando o nome ou o caminho tiver espaços.

## 10.2. Forma direta em PowerShell

Se a `.venv` já estiver preparada, é possível ignorar o iniciador:

```powershell
.\run.ps1
```

Esse comando também solicita o nome e autodetecta a única planilha válida.
Consulte os parâmetros, aliases e exemplos disponíveis:

```powershell
.\run.ps1 -Help
.\run.ps1 -h
```

## 10.3. Forma direta em Python

Não é necessário ativar a `.venv`:

```powershell
.\.venv\Scripts\python.exe -m src.main --candidate-name "SEU NOME COMPLETO"
```

O argumento também aceita `--nome`:

```powershell
.\.venv\Scripts\python.exe -m src.main --nome "SEU NOME COMPLETO"
```

Se `--input` for omitido, o programa usará a única planilha encontrada na pasta
`input`. Se houver mais de uma, ele interromperá a execução, informará os nomes
encontrados e orientará o uso de `--input`.

## 10.4. Executar sem Excel Desktop

```powershell
.\iniciar.cmd -SemPivotNativo
```

Nesse modo, o programa cria um resumo formula-driven e um gráfico comum. É útil
para ambientes sem Microsoft Excel, mas para o teste recomenda-se a Tabela
Dinâmica nativa.

## 10.5. Ver detalhes técnicos

```powershell
.\iniciar.cmd -Verbose
```

---

# 11. O que acontece durante a execução

Uma execução normal informa etapas semelhantes a:

```text
INFO | Iniciando Excel Compras Automation.
INFO | Backup criado.
INFO | Identificando abas e cabeçalhos por conteúdo.
INFO | Solicitações e políticas validadas.
INFO | Primeira versão gravada.
INFO | Abrindo o Excel Desktop em segundo plano.
INFO | Criando a Tabela Dinâmica nativa.
INFO | Criando o gráfico vinculado à Tabela Dinâmica.
INFO | Executando validações finais.
INFO | Automação concluída com sucesso.
```

Sucesso significa:

- não aparecer uma mensagem iniciada por `ERRO`;
- existir um `.xlsx` em `output`;
- existir um backup em `backup`;
- existir um `.log` em `logs`;
- o original continuar intacto.

---

# 12. Conferir o resultado

Liste o arquivo mais recente:

```powershell
Get-ChildItem .\output\*.xlsx |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 FullName, Length, LastWriteTime
```

Abra a pasta:

```powershell
Invoke-Item .\output
```

No Excel, confira:

- nome completo no arquivo;
- fórmulas nas colunas amarelas;
- Valor Total;
- Dias de Antecedência;
- Limite da Política;
- Status da Política;
- Diferença para o Limite;
- Prioridade;
- dez indicadores em `Respostas`;
- cinco solicitações prioritárias;
- justificativas e ações;
- Tabela Dinâmica;
- gráfico com título;
- destaque de solicitações críticas;
- destaque do status fora da política.

Se o Excel ainda estiver calculando, aguarde. Se necessário:

```text
Ctrl + Alt + F9
```

Depois salve.

---

# 13. Estrutura do projeto

```text
ExcelComprasAutomation/
├── .github/
│   ├── pull_request_template.md
│   └── workflows/
│       └── ci.yml
├── .vscode/
│   ├── extensions.json
│   ├── launch.json
│   └── settings.json
├── backup/
├── config/
│   ├── aliases.json
│   └── rules.json
├── docs/
├── input/
├── logs/
├── output/
├── src/
│   ├── business/
│   │   ├── policies.py
│   │   └── priorities.py
│   ├── core/
│   │   ├── engine.py
│   │   ├── exceptions.py
│   │   └── models.py
│   ├── excel/
│   │   ├── _writer_common.py
│   │   ├── detection.py
│   │   ├── excel_desktop.py
│   │   ├── validation.py
│   │   ├── workbook_writer.py
│   │   ├── writer_base.py
│   │   ├── writer_responses.py
│   │   └── writer_support.py
│   ├── services/
│   │   ├── diagnostics.py
│   │   ├── files.py
│   │   ├── logging_setup.py
│   │   └── text.py
│   ├── main.py
│   └── settings.py
├── tests/
├── CONTRIBUTING.md
├── main.py
├── pyproject.toml
├── README.md
├── requirements.txt
├── iniciar.cmd
├── run.ps1
└── setup.ps1
```

O `main.py` da raiz é apenas um atalho. A orquestração está em
`src/core/engine.py`.

## Organização interna do escritor

Desde a versão 1.3.0, `src/excel/workbook_writer.py` permanece como a fachada
estável usada pelo restante do sistema. Assim, o `AutomationEngine` continua
chamando as mesmas operações públicas enquanto a implementação fica distribuída
em módulos menores:

- `_writer_common.py`: constantes visuais e helpers compartilhados para células,
  intervalos, referências e estilos;
- `writer_base.py`: colunas derivadas, fórmulas, tabela e formatação condicional
  da base de viagens;
- `writer_support.py`: aba oculta de apoio, premissas, agregações auxiliares e
  fonte estática da Tabela Dinâmica;
- `writer_responses.py`: indicadores, cinco prioridades, área do resumo,
  fallback e gráfico.

Esses módulos são detalhes internos e ainda podem evoluir. Novas integrações
devem importar as operações de `workbook_writer.py`, evitando dependência direta
dos módulos especializados.

---

# 14. Como a detecção automática funciona

O programa não depende apenas do nome exato da aba.

Ele:

1. remove acentos, espaços duplicados e pontuação;
2. compara os nomes com aliases configuráveis;
3. examina as primeiras linhas de cada aba;
4. mede quantos cabeçalhos obrigatórios foram encontrados;
5. escolhe a melhor combinação de nome e conteúdo;
6. interrompe a execução se a confiança for insuficiente.

Exemplos que podem ser reconhecidos como o mesmo campo:

```text
Centro de Custo
Centro Custo
CC
Cost Center
```

Os aliases ficam em:

```text
config\aliases.json
```

Para aceitar um novo nome, adicione um texto à lista correspondente. Não altere
o nome canônico à esquerda.

Mais detalhes: [docs/ADAPTAR_OUTRAS_PLANILHAS.md](docs/ADAPTAR_OUTRAS_PLANILHAS.md)

---

# 15. Regras de cálculo

## Valor Total

```text
Quantidade × Valor Unitário + Taxas
```

## Dias de Antecedência

```text
Data da Viagem − Data da Solicitação
```

## Limite da Política

É localizado pelo Tipo de Serviço usando `ÍNDICE` e `CORRESP`.

## Status da Política

É `Fora` quando:

- o Valor Total supera o limite; ou
- a antecedência é menor que a antecedência mínima.

É `Revisar` quando a política não é localizada.

## Diferença para o Limite

```text
Valor Total − Limite da Política
```

## Prioridade

O score combina:

- status do cartão;
- criticidade;
- aderência à política;
- excesso de custo;
- falta de antecedência;
- valor total;
- status de reserva.

Os pesos e limites ficam em:

```text
config\rules.json
```

As premissas também são registradas em uma aba oculta chamada
`Apoio_Automacao`. As fórmulas fazem referência a essa aba para evitar números
mágicos espalhados pela planilha.

Mais detalhes: [docs/REGRAS_DE_NEGOCIO.md](docs/REGRAS_DE_NEGOCIO.md)

---

# 16. Tabela Dinâmica e gráfico

No Windows com Excel Desktop, o programa usa `pywin32` para:

1. abrir o arquivo de saída em segundo plano;
2. usar a fonte estática validada na aba oculta `Apoio_Automacao`;
3. criar um PivotCache sem depender de um recálculo global bloqueante;
4. colocar Centro de Custo nas linhas;
5. colocar Tipo de Serviço nas colunas;
6. somar Valor Total;
7. criar o gráfico sem incluir o total geral como categoria;
8. preservar as larguras e a formatação da análise gerencial;
9. marcar as fórmulas para recálculo automático ao abrir;
10. salvar e fechar o Excel.

A fonte estática contém os mesmos resultados calculados e validados pelo motor
Python. As fórmulas continuam presentes em `Base_Viagens` e `Respostas`, como
exigido no teste, mas a criação da Pivot não fica bloqueada esperando o estado
global de cálculo do Excel.

Se o Excel Desktop não puder ser iniciado, o programa preserva a entrega criando
um resumo equivalente com `SOMASES` e um gráfico comum. O log registra qual modo
foi usado.

Ao terminar uma sessão nativa, a automação tenta de forma independente fechar a
pasta de trabalho, restaurar as configurações anteriores do Excel, encerrar o
aplicativo e liberar a sessão COM. Uma falha em uma dessas etapas não impede as
demais tentativas de limpeza. Se o Excel puder ter mantido o arquivo bloqueado, a
execução é interrompida sem reabrir a saída para criar o fallback.

---

# 17. Executar os testes

Instale as dependências de desenvolvimento:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

Execute:

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

Resultado esperado na release v1.6.0: 137 testes aprovados, sem falhas ou erros.
A release v1.5.0 permanece registrada no `CHANGELOG.md` com os 126 testes
validados naquela entrega.

Os arquivos temporários dos testes são criados em `.pytest_tmp`, dentro do
próprio projeto. Isso evita erros de permissão que algumas instalações do
Windows apresentam na pasta temporária global do usuário. A pasta é descartável
e está protegida pelo `.gitignore`.

Verifique a qualidade e a formatação:

```powershell
.\.venv\Scripts\python.exe -m ruff check .
.\.venv\Scripts\python.exe -m ruff format --check .
```

Meça a cobertura:

```powershell
.\.venv\Scripts\python.exe -m pytest --cov=src --cov-report=term-missing -q
```

O projeto exige cobertura mínima de 90%. Na release v1.6.0, os 137 testes
alcançam cobertura total de 97,92%, com medição de branches habilitada. O comando
falhará se uma mudança reduzir a cobertura para menos de 90%.

## Integração contínua no GitHub

O arquivo `.github/workflows/ci.yml` executa automaticamente a validação quando:

- há um `push` nas branches `main` ou `develop`;
- um Pull Request é aberto ou atualizado para essas branches;
- a execução é iniciada manualmente pela aba **Actions** do GitHub.

O workflow usa Windows e testa as versões Python 3.11 e 3.14. Em cada versão,
ele instala as dependências, verifica a sintaxe, executa o Ruff, confere a
formatação e roda os testes com cobertura mínima de 90%.

Os runners hospedados pelo GitHub não incluem o Microsoft Excel Desktop. Por
isso, a automação COM que cria a Tabela Dinâmica nativa continua sendo validada
localmente em um computador Windows com Excel instalado. Na versão 1.6.0, as
regressões dos modos nativo e fallback processaram 40 solicitações com 0 erros
de fórmula. O modo nativo criou uma PivotTable, um gráfico e quatro regras de
formatação condicional, manteve a aba de suporte oculta e encerrou sem deixar
processos `EXCEL.EXE`. O CI valida as regras de negócio e os componentes que não
dependem da interface do Excel.

## Governança das branches

A branch `main` representa a versão publicada e é protegida. Mudanças comuns
nascem em branches curtas criadas a partir da `develop` e chegam a ela por Pull
Request. Uma publicação usa uma branch `release/vX.Y.Z`, também criada a partir
da `develop`, e chega à `main` por outro Pull Request.

Todo merge exige:

- branch atualizada em relação à branch-base do Pull Request;
- checks `Python 3.11` e `Python 3.14` aprovados;
- todas as conversas resolvidas;
- conferência do escopo e do commit efetivamente revisado.

Como este é um projeto individual, nenhuma aprovação externa é obrigatória. O
fluxo completo de desenvolvimento, validação e merge está documentado em
[CONTRIBUTING.md](CONTRIBUTING.md).

Os 137 testes automatizados verificam:

- abas renomeadas;
- colunas movidas;
- aliases;
- combinação de fatores de risco;
- ordenação determinística;
- leitura, duplicidade e ausência de políticas;
- cálculos com datas seriais e valores inválidos;
- descoberta segura da planilha de entrada;
- criação de backup e nomes de saída;
- leitura de configurações e mensagens de erro;
- criação e substituição segura dos handlers de log;
- orquestração do motor e comportamento da interface de linha de comando;
- diagnóstico somente leitura, códigos de saída e consulta de versão;
- contratos de `setup.ps1`, `run.ps1` e `iniciar.cmd`;
- escrita, fallback e validação final do arquivo;
- compatibilidade da fachada após a decomposição interna do escritor;
- ciclo de vida e limpeza resiliente da integração com o Excel Desktop usando
  simulações isoladas.

---

# 18. Próximas execuções

Python e as bibliotecas são instalados somente uma vez. O VS Code continua
opcional.

Nas próximas vezes:

1. abra a pasta do projeto;
2. mantenha uma única planilha válida em `input`;
3. dê duplo clique em `iniciar.cmd`.

O iniciador reutiliza a `.venv`, solicita o nome e mantém a janela aberta ao
final quando é aberto sem argumentos.

Se preferir usar o terminal:

```powershell
.\iniciar.cmd -NomeCompleto "SEU NOME COMPLETO"
```

Não é necessário ativar nem desativar a `.venv`. A ativação manual continua
disponível apenas para desenvolvimento. Se você optou por ativá-la, finalize com:

```powershell
deactivate
```

---

# 19. Recriar somente o ambiente virtual

Não é necessário desinstalar Python ou VS Code.

Para uma reinstalação limpa:

1. feche terminais que usam `.venv`;
2. exclua somente a pasta `.venv`;
3. abra um novo terminal;
4. execute:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

Não exclua `src`, `config`, `input`, `requirements.txt` ou o `README.md`.

---

# 20. Solução de problemas rápida

## `py` não é reconhecido

Feche e abra o terminal. Se continuar, reinstale o Python e habilite o Launcher e
o PATH.

## `pip` não é reconhecido

Use:

```powershell
.\.venv\Scripts\python.exe -m pip --version
```

## `Activate.ps1` está bloqueado durante o uso manual

O iniciador não depende da ativação. Se você estiver desenvolvendo e quiser
ativar a `.venv`, use:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\.venv\Scripts\Activate.ps1
```

## `requirements.txt` não foi encontrado

```powershell
Get-Location
Get-ChildItem
```

Volte à pasta que contém `requirements.txt`.

## `ModuleNotFoundError`

```powershell
.\.venv\Scripts\python.exe -c "import sys; print(sys.executable)"
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## `No module named src`

Execute na raiz:

```powershell
.\.venv\Scripts\python.exe -m src.main
```

## `PermissionError`

Feche os arquivos de entrada e saída no Excel e execute novamente.

## O Excel não liberou todos os recursos

Essa mensagem indica que a automação interrompeu a execução para não acessar
novamente um arquivo que ainda pode estar bloqueado. Aguarde alguns segundos para
que o Windows conclua o encerramento assíncrono e feche somente os arquivos ou as
janelas do Excel envolvidos nesta tentativa. Não encerre indiscriminadamente
processos `EXCEL.EXE` ou sessões preexistentes no Gerenciador de Tarefas. Depois,
tente novamente. Se o problema persistir, execute uma nova tentativa com:

```powershell
.\iniciar.cmd -SemPivotNativo
```

## A Tabela Dinâmica nativa não foi criada

Confirme:

- Windows;
- Excel Desktop instalado;
- `pywin32` instalado;
- nenhum aviso de segurança do Office aguardando confirmação;
- arquivo em uma pasta local.

Teste:

```powershell
.\.venv\Scripts\python.exe -c "import win32com.client; app=win32com.client.Dispatch('Excel.Application'); print(app.Version); app.Quit()"
```

Se a política corporativa bloquear automação COM, use:

```powershell
.\iniciar.cmd -SemPivotNativo
```

## O gráfico não aparece no visualizador

Abra no Microsoft Excel Desktop. Alguns visualizadores simplificados não exibem
todos os objetos.

## Erro de certificado no `pip`

Não desative a verificação de certificados. Em rede corporativa, solicite ao TI
as configurações corretas de proxy e certificado.

Consulte o guia completo:

[docs/SOLUCAO_DE_PROBLEMAS.md](docs/SOLUCAO_DE_PROBLEMAS.md)

---

# 21. Checklist final para a entrevista

- [ ] Python 3.11, 3.12, 3.13 ou 3.14 instalado.
- [ ] Projeto extraído do `.zip`.
- [ ] `iniciar.cmd` conclui a preparação da `.venv`.
- [ ] `pip check` sem conflitos.
- [ ] Uma única planilha `.xlsx` ou `.xlsm` na pasta `input`, ou `-Arquivo`
      informado.
- [ ] Planilha fechada durante a execução.
- [ ] Nome completo informado.
- [ ] Resultado aberto no Excel Desktop.
- [ ] Fórmulas preservadas.
- [ ] Dez indicadores conferidos.
- [ ] Cinco prioridades justificadas.
- [ ] Tabela Dinâmica e gráfico conferidos.
- [ ] Original preservado.
- [ ] Arquivo final renomeado com o nome completo.
- [ ] GitHub Actions aprovado no Pull Request.

Para estudar ou modificar o projeto, instale opcionalmente o VS Code, Python e
Pylance e selecione o interpretador da `.venv`.

---

# Documentação adicional

- [Como contribuir](CONTRIBUTING.md)
- [Arquitetura](docs/ARQUITETURA.md)
- [Regras de negócio](docs/REGRAS_DE_NEGOCIO.md)
- [Adaptação para outras planilhas](docs/ADAPTAR_OUTRAS_PLANILHAS.md)
- [Solução de problemas](docs/SOLUCAO_DE_PROBLEMAS.md)
- [Histórico de versões](CHANGELOG.md)
