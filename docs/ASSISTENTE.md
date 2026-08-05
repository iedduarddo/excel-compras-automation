# Assistente local de planilhas

O assistente adiciona uma pasta monitorada ao motor existente. Nesta primeira
etapa, ele reconhece formatos compatíveis, escolhe adaptadores salvos e executa
somente comandos conhecidos. O arquivo original não é sobrescrito.

## Preparar as pastas

```powershell
.\run.ps1 -Assistente -PrepararPastas
```

A estrutura será criada em `assistente_planilhas`:

```text
assistente_planilhas/
├── entrada/
├── saida/
├── backup/
├── logs/
├── revisao/
├── adaptadores/
├── comandos/
│   ├── pendentes/
│   ├── concluidos/
│   └── falhas/
└── config.json
```

Preencha `candidate_name` em `config.json` para não repetir o nome em todos os
comandos. `use_native_pivot` controla o uso do Excel Desktop.

## Comandos escritos

Execução direta:

```powershell
.\run.ps1 -Assistente -Comando 'reconhecer todas'
.\run.ps1 -Assistente -Comando 'diagnosticar todas'
.\run.ps1 -Assistente `
    -Comando 'processar todas nome="Maria Aparecida" sem excel'
.\run.ps1 -Assistente `
    -Comando 'processar arquivo="compras.xlsx" nome="Maria Aparecida"'
```

Também é possível criar arquivos `.txt` em `comandos\pendentes` e executar:

```powershell
.\run.ps1 -Assistente
```

Cada comando será movido para `concluidos` ou `falhas`, acompanhado de um JSON
com o resultado. Para manter a pasta sob observação:

```powershell
.\run.ps1 -Assistente -Monitorar
```

Interrompa o monitor com `Ctrl+C`.

## Comando por voz no Windows

O modo de voz abre uma janela de confirmação e aciona a Digitação por Voz do
Windows (`Win+H`):

```powershell
.\run.ps1 -Assistente -Voz
```

Fale um dos mesmos comandos aceitos por escrito, por exemplo `ajuda`,
`reconhecer todas`, `diagnosticar todas` ou `processar todas`. Revise o texto na
janela e clique em `Usar comando`; somente o texto confirmado é entregue ao
interpretador seguro.

O recurso requer microfone autorizado, Windows com Digitação por Voz e o teclado
`Português (Brasil)` instalado. O tratamento do áudio pertence ao Windows e
segue as configurações de privacidade do sistema; a automação não exige chave de
API. A voz é opcional: a fila `.txt` e `-Comando` continuam disponíveis.

## Reconhecimento e revisão

O assistente tenta primeiro o mapeamento padrão e depois os perfis `.json` da
pasta `adaptadores`. Se nenhum reconhecer a estrutura, nenhuma automação é
executada: um relatório com abas e possíveis cabeçalhos é salvo em `revisao`.

Esse relatório é a entrada para criar um novo perfil conforme
`docs/ADAPTADORES.md`. Depois que o perfil for salvo, as próximas planilhas do
mesmo formato poderão ser reconhecidas automaticamente.

## Limites desta etapa

Os comandos disponíveis são `ajuda`, `reconhecer`, `diagnosticar` e
`processar`. Pedidos livres como apagar dados, enviar e-mail ou inventar
fórmulas são recusados, tanto por voz quanto por escrita. A voz apenas transcreve
para o mesmo interpretador seguro; ela não amplia as operações permitidas.
