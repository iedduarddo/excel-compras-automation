# Adaptadores de planilhas

Existem dois tipos de adaptador, com responsabilidades diferentes.

## Adaptador do fluxo de Compras

Esse perfil acrescenta nomes externos ao vocabulário de abas, colunas e
indicadores já conhecido pela automação de Compras. Ele não substitui a
configuração padrão.

```json
{
  "name": "agencia-exemplo",
  "aliases": {
    "sheets": {
      "base": ["Pedidos da Agência"],
      "policies": ["Parâmetros da Agência"],
      "responses": ["Indicadores da Agência"]
    },
    "base_columns": {
      "request_id": ["Número do Pedido"],
      "request_date": ["Data de Abertura"]
    },
    "policy_columns": {
      "limit_value": ["Teto Autorizado"]
    },
    "response_columns": {
      "indicator": ["Métrica Calculada"]
    }
  }
}
```

Os grupos aceitos são `sheets`, `base_columns`, `policy_columns`,
`response_columns` e `indicator_labels`. Cada chave interna precisa existir em
`config/aliases.json`.

Valide antes de processar:

```powershell
.\iniciar.cmd `
    -Diagnostico `
    -Arquivo "C:\Planilhas\entrada.xlsx" `
    -Adaptador "C:\Mapeamentos\agencia-exemplo.json"
```

## Adaptador universal gerado

Para uma estrutura ainda desconhecida, coloque o arquivo em
`assistente_planilhas\entrada` e peça:

```powershell
.\run.ps1 -Assistente `
    -Comando 'gerar adaptador arquivo="origem-nova.xlsx"'
```

O assistente cria primeiro uma prévia. Somente depois de `confirmar plano` o
JSON é gravado em `assistente_planilhas\adaptadores\universais`.

Esse perfil registra a linha de cabeçalho, o nome original, a posição e o tipo
observado de cada coluna. Ele serve para auditar e reaproveitar o mapeamento em
ações universais; não é carregado como alias do fluxo específico de Compras.

Nenhum dos dois formatos altera o arquivo original.
