# Arquitetura

## Objetivo

A arquitetura separa infraestrutura de Excel, regras de negócio e orquestração.
Isso permite reutilizar os detectores e serviços em projetos futuros.

## Fluxo

```text
src.main
   |
   +--> --version
   |       |
   |       +--> exibe a versão e encerra
   |
   +--> --diagnostico / --diagnostic
   |       |
   |       v
   |   run_diagnostics
   |       |
   |       +--> verifica ambiente e estrutura da entrada
   |       +--> não cria backup, saída ou log
   |
   +--> execução normal
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

É a fachada estável da escrita do Excel. O `AutomationEngine` continua
importando dela as operações públicas para abrir a origem, preparar a base,
criar a aba de apoio, escrever fórmulas e respostas, gerar o fallback, aplicar
formatação e configurar o recálculo.

Desde a versão 1.3.0, a implementação interna está decomposta nos módulos abaixo
sem alterar esse contrato público. Código externo ao pacote `src.excel` deve
depender da fachada, e não dos módulos especializados.

### `src/excel/_writer_common.py`

Centraliza constantes visuais e helpers compartilhados para endereços de
células, intervalos absolutos, nomes de abas e estilos. O prefixo `_` indica que
é um detalhe interno, sujeito a mudanças durante a evolução da arquitetura.

### `src/excel/writer_base.py`

Cuida da estrutura da base de viagens: cria ou reaproveita colunas derivadas,
estende a tabela, escreve as fórmulas auditáveis e aplica formatação condicional.

### `src/excel/writer_support.py`

Cria novamente a aba oculta `Apoio_Automacao`, registra pesos e limites, monta
agregações auxiliares e prepara a fonte estática usada pela PivotTable nativa.

### `src/excel/writer_responses.py`

Preenche os indicadores, escreve e formata as solicitações prioritárias, reserva
a área da Tabela Dinâmica e gera o resumo com gráfico usado como fallback.

### `src/excel/excel_desktop.py`

Isola a dependência do Windows e do Excel. Cria a PivotTable nativa e o gráfico
a partir de uma fonte estática validada, preserva a formatação, configura o
recálculo automático para a próxima abertura e salva. A criação da Pivot não
depende de `CalculateFullRebuild` nem do estado global de cálculo do Excel.

### `src/excel/validation.py`

Verifica fórmulas, erros, total conciliado, formatação, gráfico e partes XML da
PivotTable.

### `src/services`

Contém operações reutilizáveis de arquivos, logging e texto. O módulo
`diagnostics.py` verifica a versão do Python, as configurações, a entrada, a
estrutura da planilha e a disponibilidade da integração nativa sem iniciar o
Excel Desktop nem gerar artefatos operacionais.

## Decisões

- O original nunca é salvo.
- A entrada é lida com `data_only=False` para preservar fórmulas.
- Fórmulas usam referências detectadas, e não colunas fixas.
- Intervalos são limitados às linhas usadas, evitando referências como `P:P`.
- Pesos ficam fora do código.
- O score é reproduzido em Python e no Excel.
- O fallback mantém o projeto executável fora do Excel Desktop.
- O Engine depende da fachada `workbook_writer.py`; os módulos `writer_*` ficam
  encapsulados como detalhes da camada de Excel.

## Evolução sugerida

1. extrair um pacote genérico de detecção;
2. definir schemas de negócio por plugin;
3. incluir testes de integração com Excel em uma máquina dedicada;
4. criar interface gráfica;
5. empacotar com instalador;
6. registrar histórico em SQLite.
