# Excel Compras Automation — pacote portátil para Windows 11 x64

Este pacote executa sem instalar Python, sem criar `.venv` e sem baixar
dependências da internet.

## Como usar

1. Extraia **todo** o conteúdo do ZIP para uma pasta gravável, como
   `Documentos\ExcelComprasAutomation`.
2. Não execute o programa diretamente de dentro do ZIP e não mova apenas o
   `.exe`: a pasta `_internal` precisa permanecer ao lado dele.
3. Copie uma única planilha `.xlsx` ou `.xlsm` para a pasta `input`.
4. Dê duplo clique em `iniciar.cmd`.
5. Informe nome e sobrenome quando o terminal solicitar.

O resultado será criado em `output`; o original será preservado e uma cópia
será guardada em `backup`. Os registros técnicos ficam em `logs`.

## Diagnóstico e opções

Abra o PowerShell na pasta extraída e use:

```powershell
.\iniciar.cmd -Diagnostico
.\iniciar.cmd -Version
.\iniciar.cmd -Help
.\iniciar.cmd -NomeCompleto "NOME SOBRENOME"
.\iniciar.cmd -NomeCompleto "NOME SOBRENOME" -SemPivotNativo
.\iniciar.cmd `
    -Arquivo ".\input\minha_planilha.xlsx" `
    -NomeCompleto "NOME SOBRENOME"
```

O Microsoft Excel Desktop é opcional. Quando ele não está disponível, a
aplicação produz automaticamente o resumo compatível por fórmulas. A opção
`-SemPivotNativo` força esse modo.

## Configuração

Os arquivos `config\aliases.json` e `config\rules.json` permanecem externos e
editáveis. Preserve a estrutura JSON e faça uma cópia antes de alterá-los.

Para uma origem com nomes diferentes de abas ou colunas, mantenha a
configuração padrão intacta e informe um perfil separado:

```powershell
.\iniciar.cmd `
    -Diagnostico `
    -Arquivo ".\input\planilha_cliente.xlsx" `
    -Adaptador ".\config\cliente.json"
```

Depois do diagnóstico aprovado, repita o parâmetro na execução normal ou em
lote. O formato está documentado em `docs\ADAPTADORES.md`.

## Pasta monitorada

O pacote também pode criar uma central isolada para entradas e comandos:

```powershell
.\iniciar.cmd -Assistente -PrepararPastas
.\iniciar.cmd -Assistente -Comando 'diagnosticar todas'
.\iniciar.cmd -Assistente -Voz
.\iniciar.cmd -Assistente -Monitorar
```

Copie as planilhas para `assistente_planilhas\entrada`. Consulte
`docs\ASSISTENTE.md` para configurar o nome, usar a fila e entender quando um
arquivo é enviado para revisão.

Para estruturas genéricas, crie uma prévia e confirme somente depois de
revisá-la:

```powershell
.\iniciar.cmd -Assistente `
    -Comando 'limpar e resumir arquivo="clientes.xlsx"'
.\iniciar.cmd -Assistente `
    -Comando 'confirmar plano="IDENTIFICADOR"'
```

O modo `-Voz` abre uma janela e aciona a Digitação por Voz do Windows (`Win+H`).
Revise o texto e confirme antes da execução. A automação não exige chave de API,
e a transcrição continua limitada aos mesmos comandos seguros da escrita.

## Segurança do download

O executável ainda não possui assinatura de código. Por isso, o Windows
SmartScreen pode exibir um aviso mesmo quando o arquivo foi baixado da página
oficial do projeto.

Antes de extrair, abra o PowerShell na pasta de downloads. Compare o SHA-256 do
ZIP com o arquivo `.sha256` publicado junto da versão:

```powershell
Get-FileHash `
    .\ExcelComprasAutomation-vX.Y.Z-windows-x64.zip `
    -Algorithm SHA256
```

O valor exibido deve ser idêntico ao primeiro campo do arquivo
`ExcelComprasAutomation-vX.Y.Z-windows-x64.zip.sha256`.

Ao atualizar, extraia a nova versão em outra pasta. Migre somente suas planilhas
e alterações deliberadas em `config`; não sobreponha `_internal` com arquivos
de uma versão antiga.

## Solução rápida de problemas

- Mantenha `ExcelComprasAutomation.exe`, `_internal`, `config`, `run.ps1` e
  `iniciar.cmd` na mesma pasta.
- Extraia o pacote para uma pasta em que seu usuário possa criar arquivos.
- Feche a planilha de entrada antes da execução.
- Deixe somente uma planilha na pasta `input` ou informe `-Arquivo`.
- Execute `.\iniciar.cmd -Diagnostico` para verificar a entrada e a
  configuração sem criar backup, saída ou log.
