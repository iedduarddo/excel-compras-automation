# Regras de negócio

## Política

Cada serviço possui:

- limite de valor;
- antecedência mínima.

O status é:

```text
Revisar
```

quando o serviço não existe na tabela de políticas.

É:

```text
Fora
```

quando pelo menos uma condição é verdadeira:

```text
Valor Total > Limite da Política
Dias de Antecedência < Antecedência Mínima
```

Nos demais casos:

```text
OK
```

## Score de prioridade

Pesos padrão:

| Fator | Pontos |
|---|---:|
| Cartão divergente | 40 |
| Cartão pendente | 30 |
| Outro problema de cartão | 25 |
| Emergencial | 35 |
| Executivo | 25 |
| Fora da política | 30 |
| Política não localizada | 40 |
| Reserva pendente | 8 |
| Remarcação | 5 |
| Excesso de custo | até 25 |
| Falta de antecedência | até 20 |
| Valor total | até 10 |

Classificação:

| Score | Prioridade |
|---:|---|
| 75 ou mais | Crítica |
| 25 a 74,99 | Alta |
| abaixo de 25 | Normal |

Os valores podem ser alterados em `config/rules.json`.

## Componentes proporcionais

### Excesso de custo

```text
min(
    peso máximo,
    Diferença Positiva / Limite × peso máximo
)
```

### Falta de antecedência

```text
min(
    peso máximo,
    (Mínimo − Dias Reais) / Mínimo × peso máximo
)
```

### Valor total

```text
Valor da Solicitação / Maior Valor da Base × peso máximo
```

## Cinco solicitações

A ordenação usa:

1. score decrescente;
2. Valor Total decrescente;
3. ID crescente como desempate.

Isso torna o resultado reproduzível.

## Indicadores

1. Valor total das viagens.
2. Valor total fora da política.
3. Quantidade fora da política.
4. Percentual fora da política.
5. Valor pendente ou divergente no cartão.
6. Quantidade emergencial.
7. Centro de custo com maior despesa.
8. Fornecedor com maior valor.
9. Economia potencial.
10. Ticket médio.

A economia potencial considera somente diferenças positivas. Uma solicitação
fora da política apenas por antecedência, mas abaixo do limite financeiro, não
gera uma economia negativa.

## Observação sobre `Limite por Viajante`

Nesta planilha, `Limite Política` é o valor localizado diretamente pelo Tipo de
Serviço, conforme o layout e a fórmula solicitados no teste. Se uma empresa
definir que o limite deve ser multiplicado por quantidade, a regra deverá ser
alterada explicitamente e documentada.

