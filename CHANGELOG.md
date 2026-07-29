# Histórico de versões

Todas as mudanças relevantes deste projeto serão registradas neste arquivo.
O formato segue a ideia do [Keep a Changelog](https://keepachangelog.com/),
com versionamento semântico.

## Em desenvolvimento

### Alterado

- fluxo de contribuição alinhado ao uso de branches curtas para `develop` e
  branches de release para `main`;
- template de Pull Request atualizado para considerar a branch-base correta;
- arquitetura atualizada com os caminhos de versão, diagnóstico e execução
  normal, além do estado concluído da decomposição do escritor;
- documentação da governança e da regressão local atualizada para a versão
  1.6.0.
- contrato do Engine alinhado ao retorno real da criação nativa, removendo uma
  segunda tentativa de recálculo que não possuía caminho executável.

### Planejado

- avaliar um runner Windows próprio para testes de integração com Excel Desktop;
- preparar empacotamento e interface para usuários não técnicos.

## 1.6.0 — 2026-07-29

### Adicionado

- iniciador `iniciar.cmd` para preparar o ambiente quando necessário, executar
  `run.ps1`, repassar argumentos e preservar o código de saída;
- ajuda por `-Help` ou `-h` com opções e exemplos de uso do `run.ps1`;
- testes de contrato dos scripts Windows sem instalar pacotes nem iniciar o
  Excel Desktop.

### Alterado

- `setup.ps1` agora aceita somente Python 3.11 a 3.14, valida cada código de
  saída, cria as pastas operacionais e reutiliza com segurança uma `.venv`
  compatível;
- `run.ps1` passou a autodetectar a única planilha `.xlsx` ou `.xlsm` da pasta
  `input`, enviando `--input` somente quando `-Arquivo` é informado;
- onboarding recomendado por duplo clique em `iniciar.cmd`, sem exigir VS Code
  ou ativação manual do ambiente virtual;
- janela do iniciador mantida aberta ao final quando ele é chamado sem
  argumentos, preservando a execução não interativa quando há parâmetros.

### Validado

- 137 testes aprovados, com cobertura total de 97,92%;
- contratos de versão do Python, idempotência, diretórios operacionais,
  autodetecção da entrada, switches públicos e códigos de saída aprovados.

## 1.5.0 — 2026-07-29

### Adicionado

- erro específico `ExcelDesktopCleanupError` para indicar que o Excel pode não
  ter liberado a pasta de trabalho ou o processo;
- testes de contrato com objetos COM simulados para inicialização parcial,
  criação da PivotTable, recálculo e falhas independentes de limpeza.

### Corrigido

- fechamento da pasta de trabalho, restauração das configurações do Excel,
  encerramento do aplicativo e liberação da sessão COM executados de forma
  independente, mesmo quando uma etapa anterior falha;
- preservação do erro principal, com falhas adicionais de limpeza registradas
  sem impedir as tentativas seguintes;
- interrupção segura do Engine após uma falha crítica de limpeza, evitando
  reabrir uma saída que ainda pode estar bloqueada pelo Excel.

### Validado

- 126 testes aprovados, com cobertura total de 97,92%;
- módulo `src/excel/excel_desktop.py` com 240 statements sem linhas ausentes e
  99% de cobertura considerando branches;
- regressões reais dos modos nativo e fallback aprovadas com 40 solicitações e
  0 erros de fórmula;
- modo nativo aprovado com uma PivotTable, um gráfico, quatro regras de
  formatação condicional, aba de suporte oculta, arquivo original preservado e
  nenhuma instância residual de `EXCEL.EXE`.

## 1.4.0 — 2026-07-28

### Adicionado

- diagnóstico somente leitura por `--diagnostico` ou `--diagnostic`, sem
  solicitação de nome e sem criação de backup, saída ou log;
- verificações da versão do Python, configurações, planilha de entrada, abas
  obrigatórias e disponibilidade da integração pywin32;
- consulta da versão instalada por `--version`;
- opções equivalentes `-Diagnostico`, `-Diagnostic` e `-Version` no `run.ps1`,
  com propagação do código de saída.

### Validado

- 108 testes aprovados, com cobertura total de 91,76%;
- Ruff aprovado para análise estática e formatação;
- diagnóstico fecha a planilha inspecionada e não inicia o Excel Desktop.

## 1.3.0 — 2026-07-28

### Adicionado

- módulos especializados `_writer_common.py`, `writer_base.py`,
  `writer_support.py` e `writer_responses.py`;
- testes de compatibilidade da fachada e de caracterização do pipeline de escrita.

### Alterado

- `workbook_writer.py` transformado em fachada estável para preservar o contrato
  usado pelo Engine enquanto a implementação fica dividida por responsabilidade.

### Validado

- 92 testes aprovados, com cobertura total de 91,73%;
- CI aprovado no Windows com Python 3.11 e Python 3.14;
- regressões dos modos nativo e fallback aprovadas com 40 solicitações e 0 erros
  de fórmula;
- regras de negócio, fórmulas e resultado final preservados após a decomposição.

## 1.2.0 — 2026-07-28

### Adicionado

- suíte ampliada de regressão para linha de comando, engine, integração simulada
  com o Excel Desktop, detecção, prioridades, validação e escrita da planilha;
- testes dos caminhos de sucesso e falha da validação do arquivo final;
- testes de indicadores movidos, limites de varredura, colisões de abas e
  estruturas incompletas.

### Alterado

- cobertura mínima obrigatória elevada de 34% para 90%.

### Corrigido

- fechamento garantido dos workbooks abertos durante a validação final;
- mensagens claras para valores cacheados inválidos;
- validação explícita da existência e ocultação da aba `Apoio_Automacao`;
- tratamento da interrupção pelo usuário durante a solicitação do nome.

### Validado

- 90 testes aprovados, com cobertura total de 91,58%;
- regressão completa aprovada no Windows com Excel Desktop: 40 solicitações,
  nenhuma fórmula inválida, Tabela Dinâmica nativa, gráfico, quatro regras de
  formatação condicional e planilha de suporte oculta;
- regras de negócio, fórmulas e resultado final da planilha preservados.

## 1.1.0 — 2026-07-28

### Adicionado

- Ruff para análise estática, ordenação de imports e formatação;
- pytest-cov para medição da cobertura dos testes;
- configuração reproduzível das ferramentas no `pyproject.toml`;
- testes de políticas, cálculos, arquivos, configurações e logging;
- cenários de erro para datas, números, entradas ambíguas e JSON inválido;
- cobertura mínima inicial de 34% para impedir regressões;
- integração contínua no GitHub Actions para Windows e Python 3.11/3.14;
- proteção da branch `main` com Pull Request, checks e resolução de conversas
  obrigatórios, além de bloqueio de force push e exclusão;
- guia de contribuição e modelo padronizado de Pull Request;
- teste de consistência entre a versão do pacote e o `pyproject.toml`.

### Corrigido

- imports não utilizados e fora da ordem;
- tratamento silencioso de erro durante a restauração do Excel;
- criação e fechamento do arquivo temporário de reparo OOXML;
- geração de timestamps com informação de fuso horário;
- fechamento explícito dos handlers antigos ao reconfigurar o logger.

## 1.0.0 — 2026-07-27

### Adicionado

- detecção automática de abas, cabeçalhos e colunas;
- leitura parametrizada das políticas de fornecedores;
- fórmulas de Valor Total, antecedência, política, diferença e prioridade;
- indicadores e seleção das cinco solicitações mais urgentes;
- Tabela Dinâmica nativa com gráfico no Excel Desktop;
- resumo compatível quando o Excel Desktop não está disponível;
- formatação condicional para exceções e prioridades;
- backup automático, logs e validação final;
- configuração externa de aliases, pesos e limites;
- testes básicos e execução por linha de comando.

### Validado

- 40 solicitações e 4 políticas processadas;
- 280 fórmulas da base e 10 fórmulas de indicadores verificadas;
- nenhuma fórmula inválida na entrega;
- PivotTable nativa e gráfico presentes;
- arquivo original preservado.
