# Assistente local de planilhas

O assistente recebe planilhas em uma pasta isolada e aceita os mesmos pedidos
por escrita ou voz. O processamento específico de Compras continua disponível,
mas pedidos universais usam um perfil genérico e não dependem de nomes fixos de
abas ou colunas.

O arquivo original nunca é sobrescrito.

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
│   └── planos/
├── adaptadores/
│   └── universais/
├── comandos/
│   ├── pendentes/
│   ├── concluidos/
│   └── falhas/
└── config.json
```

## Pedidos por escrita ou voz

Exemplos:

```powershell
.\run.ps1 -Assistente -Comando 'reconhecer todas'
.\run.ps1 -Assistente -Comando 'diagnosticar todas'
.\run.ps1 -Assistente `
    -Comando 'limpar e organizar arquivo="clientes.xlsx" remover duplicados'
.\run.ps1 -Assistente `
    -Comando 'calcular e resumir arquivo="vendas.xlsx" coluna="Valor"'
.\run.ps1 -Assistente `
    -Comando 'criar relatório arquivo="vendas.xlsx"'
.\run.ps1 -Assistente `
    -Comando 'gerar adaptador arquivo="origem-nova.xlsx"'
.\run.ps1 -Assistente -Voz
```

O modo de voz abre a Digitação por Voz do Windows (`Win+H`). Revise a frase e
clique em `Usar comando`. A transcrição passa exatamente pelo mesmo
interpretador seguro usado na escrita; a voz não libera ações adicionais.

## Prévia e confirmação obrigatórias

Pedidos universais não alteram arquivos imediatamente. A primeira execução:

1. reconhece tabelas, cabeçalhos e tipos;
2. calcula SHA-256 da entrada;
3. grava a prévia `.md` e o plano `.json` em `revisao\planos`;
4. exibe o identificador do plano e para.

Depois de revisar a prévia, confirme explicitamente:

```powershell
.\run.ps1 -Assistente -Comando 'confirmar plano="IDENTIFICADOR"'
```

Ou cancele sem gerar arquivos:

```powershell
.\run.ps1 -Assistente -Comando 'cancelar plano="IDENTIFICADOR"'
```

Se a planilha mudar entre a prévia e a confirmação, o SHA-256 não combina e a
execução é recusada. Um novo plano deve ser criado.

## Ações universais permitidas

- `limpar`: normaliza espaços e células textuais vazias; duplicidades só são
  removidas quando o pedido diz `remover duplicados`;
- `organizar`: aplica filtros, congela cabeçalhos e ajusta larguras; uma coluna
  pode ser indicada com `ordenar por="Coluna"`;
- `calcular`: cria uma aba com contagem, soma, média, mínimo e máximo das
  colunas numéricas, opcionalmente limitada por `coluna="Valor"`;
- `resumir`: cria uma aba com quantidade de linhas, campos, vazios e duplicados;
- `criar relatório`: cria uma aba executiva com completude e tipos observados;
- `gerar adaptador`: cria um JSON universal com abas, cabeçalhos, posições e
  tipos em `adaptadores\universais`.

Várias ações podem ser combinadas na mesma frase. Para eliminar ambiguidade em
nomes repetidos, informe também `aba="Nome da aba"`.

## Fila e monitoramento

Arquivos `.txt` colocados em `comandos\pendentes` usam o mesmo fluxo. Execute
uma vez com `-Assistente` ou mantenha o monitor ativo:

```powershell
.\run.ps1 -Assistente -Monitorar
```

Cada comando é arquivado em `concluidos` ou `falhas` com um resultado JSON.

## Limites de segurança

O interpretador não executa Python, macros, comandos de sistema ou fórmulas
livres. Pedidos para apagar originais, sobrescrever arquivos ou enviar dados
são recusados. O resultado é sempre uma nova cópia na pasta `saida`, acompanhada
de backup quando a planilha é transformada.
