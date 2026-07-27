# Adaptar para outras planilhas

## Mudança apenas de nomes

Se a estrutura lógica for a mesma, normalmente basta editar
`config/aliases.json`.

Exemplo:

```json
"cost_center": [
  "Centro de Custo",
  "Centro Custo",
  "CC",
  "Cost Center",
  "Unidade Orçamentária"
]
```

Não mude `cost_center`. Adicione somente um novo alias à lista.

## Mudança de posição

Nenhuma configuração é necessária. O detector examina a linha de cabeçalho e
guarda o número real de cada coluna.

## Cabeçalho em outra linha

O detector examina as primeiras 30 linhas. Um título, logotipo ou texto antes da
tabela não impede a identificação.

## Aba renomeada

Adicione o nome em `sheets`:

```json
"base": [
  "Base_Viagens",
  "Dados de Deslocamentos"
]
```

O conteúdo também participa da decisão, portanto o nome exato não é o único
critério.

## Nova regra de prioridade

Pesos podem ser alterados em `config/rules.json`.

Uma regra completamente nova exige dois ajustes:

1. cálculo em `src/business/priorities.py`;
2. fórmula correspondente em `src/excel/workbook_writer.py`.

As duas implementações devem permanecer equivalentes.

## Outro domínio

Para estoque, vendas, RH ou financeiro, reutilize:

- `services/text.py`;
- `services/files.py`;
- `services/logging_setup.py`;
- padrão de `detection.py`;
- integração com Excel;
- validação estrutural.

Crie novos módulos de negócio para os campos e as regras daquele domínio.

## Quando o detector deve parar

Ele deve falhar quando:

- faltam campos obrigatórios;
- duas áreas são indistinguíveis;
- um cabeçalho é ambíguo;
- o tipo de dado é inválido.

Uma automação corporativa não deve adivinhar silenciosamente um campo financeiro
crítico.

