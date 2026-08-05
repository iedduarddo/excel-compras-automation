# Adaptadores de planilhas

Um adaptador acrescenta nomes externos ao vocabulário já conhecido pela
automação. Ele não substitui as configurações padrão e não altera fórmulas,
regras de negócio ou arquivos originais.

## Formato

Crie um JSON como o exemplo abaixo:

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
      "request_date": ["Data de Abertura"],
      "travel_date": ["Data de Utilização"]
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
`response_columns` e `indicator_labels`. Cada chave interna precisa existir no
`config/aliases.json`; isso impede erros silenciosos de digitação.

## Validar antes de processar

```powershell
.\iniciar.cmd `
    -Diagnostico `
    -Arquivo "C:\Planilhas\entrada.xlsx" `
    -Adaptador "C:\Mapeamentos\agencia-exemplo.json"
```

## Executar

```powershell
.\iniciar.cmd `
    -NomeCompleto "NOME SOBRENOME" `
    -Arquivo "C:\Planilhas\entrada.xlsx" `
    -Adaptador "C:\Mapeamentos\agencia-exemplo.json"
```

O mesmo adaptador pode ser usado com `-Lote`. Nesta primeira versão, o perfil
adapta nomes de abas, colunas e indicadores. Transformações de valores ou a
criação de áreas ausentes devem ser implementadas em uma etapa específica.
