# Arquitetura

## Objetivo

A arquitetura separa infraestrutura de Excel, regras de negócio e orquestração.
Isso permite reutilizar os detectores e serviços em projetos futuros.

## Fluxo

```text
src.main
   |
   v
AutomationEngine
   |
   +--> valida entrada e nome
   +--> cria backup e log
   +--> detecta abas e colunas
   +--> lê políticas
   +--> calcula e pontua solicitações
   +--> escreve fórmulas e respostas
   +--> cria PivotTable/gráfico
   +--> valida o arquivo
   |
   v
arquivo em output
```

## Responsabilidades

### `src/main.py`

Interpreta os argumentos do terminal e apresenta o resultado. Não conhece
detalhes de Excel.

### `src/core/engine.py`

Coordena a sequência. Ele não implementa diretamente fórmulas, detecção ou
pontuação.

### `src/excel/detection.py`

Normaliza textos, encontra linhas de cabeçalho e identifica colunas sem depender
de posição fixa.

### `src/business/policies.py`

Lê políticas e calcula em memória os resultados que serão usados na validação e
na seleção das cinco prioridades.

### `src/business/priorities.py`

Aplica os pesos de `config/rules.json`, cria justificativas e ordena riscos.

### `src/excel/workbook_writer.py`

Escreve fórmulas, formatação condicional, indicadores, análise das cinco
prioridades e o resumo de compatibilidade.

### `src/excel/excel_desktop.py`

Isola a dependência do Windows e do Excel. Cria a PivotTable nativa e o gráfico
a partir de uma fonte estática validada, preserva a formatação, configura o
recálculo automático para a próxima abertura e salva. A criação da Pivot não
depende de `CalculateFullRebuild` nem do estado global de cálculo do Excel.

### `src/excel/validation.py`

Verifica fórmulas, erros, total conciliado, formatação, gráfico e partes XML da
PivotTable.

### `src/services`

Contém operações reutilizáveis de arquivo, log e texto.

## Decisões

- O original nunca é salvo.
- A entrada é lida com `data_only=False` para preservar fórmulas.
- Fórmulas usam referências detectadas, e não colunas fixas.
- Intervalos são limitados às linhas usadas, evitando referências como `P:P`.
- Pesos ficam fora do código.
- O score é reproduzido em Python e no Excel.
- O fallback mantém o projeto executável fora do Excel Desktop.

## Evolução sugerida

1. extrair um pacote genérico de detecção;
2. definir schemas de negócio por plugin;
3. incluir testes de integração com Excel em uma máquina dedicada;
4. criar interface gráfica;
5. empacotar com instalador;
6. registrar histórico em SQLite.
