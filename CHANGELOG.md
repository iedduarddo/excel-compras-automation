# Histórico de versões

Todas as mudanças relevantes deste projeto serão registradas neste arquivo.
O formato segue a ideia do [Keep a Changelog](https://keepachangelog.com/),
com versionamento semântico.

## Em desenvolvimento

### Adicionado

- reconhecimento genérico de tabelas, cabeçalhos, tipos, vazios e duplicidades
  em planilhas que ainda não possuem mapeamento específico;
- pedidos universais por voz ou escrita para limpar, organizar, calcular,
  resumir, criar relatório e gerar adaptador;
- plano persistente com prévia legível, SHA-256 da origem e confirmação
  obrigatória antes de criar qualquer arquivo transformado;
- adaptadores universais gerados com abas, cabeçalhos, posições e tipos, sem
  misturá-los aos aliases do fluxo específico de Compras.

### Segurança

- originais nunca são sobrescritos e toda transformação gera nova saída e
  backup;
- planos são invalidados quando a origem muda após a prévia;
- pedidos para apagar, sobrescrever ou enviar dados continuam recusados.

### Planejado

- avaliar um runner Windows próprio para testes de integração com Excel Desktop;
- avaliar assinatura de código para reduzir avisos do Windows SmartScreen;
- estudar uma interface gráfica sem alterar o fluxo de linha de comando.

## 1.10.0 - 2026-08-05

### Adicionado

- entrada `-Voz` do assistente usando a Digitacao por Voz do Windows, com
  revisao humana e limitada aos mesmos comandos seguros da escrita;
- janela para falar, revisar a transcricao e confirmar antes da execucao;
- testes da transcricao confirmada, erros da interface do Windows e
  integracao com a linha de comando.

### Validado

- regressao real em pt-BR confirmou o comando `ajuda` antes da execucao;
- 217 testes aprovados no Windows, com cobertura total de 93,34%;
- CI aprovado em Python 3.11 e Python 3.14;
- pacote portatil Windows x64 aprovado no commit integrado.

## 1.9.1 - 2026-08-05

### Corrigido

- consumo atomico dos comandos pendentes para impedir que dois monitores
  processem ou movam o mesmo arquivo simultaneamente.

### Validado

- teste concorrente deterministico confirmou um unico processamento;
- 203 testes aprovados no Windows, com cobertura total de 93,13%;
- CI aprovado em Python 3.11 e Python 3.14;
- pacote portatil Windows x64 aprovado no commit integrado.

## 1.9.0 - 2026-08-04

### Adicionado

- perfis JSON de adaptador para acrescentar nomes específicos de abas,
  colunas e indicadores sem alterar o mapeamento padrão;
- opção `--adaptador` na interface Python e `-Adaptador` nos iniciadores
  Windows, disponível também no diagnóstico e no processamento em lote;
- validação estrita dos grupos e campos canônicos declarados pelo adaptador.
- pasta `assistente_planilhas` com entrada, saída, backup, revisão, adaptadores,
  logs e fila persistente de comandos;
- comandos escritos para reconhecer, diagnosticar e processar uma ou várias
  planilhas, com modo de monitoramento local;
- encaminhamento seguro de formatos desconhecidos para revisão, sem modificar
  o arquivo original.

### Validado

- 202 testes aprovados no Windows, com cobertura total de 93,18%;
- CI aprovado em Python 3.11 e Python 3.14;
- pacote portatil Windows x64 aprovado no commit integrado;
- regressao real aprovada com reconhecimento automatico, diagnostico e
  processamento por comando escrito, preservando a planilha original.

## 1.8.0 — 2026-08-04

### Adicionado

- modo `--lote` na interface Python e `-Lote` nos scripts Windows para
  processar todas as planilhas válidas da pasta `input`;
- resumo final com sucessos, falhas, saídas e motivos por entrada;
- testes da continuidade após falha esperada e do aborto seguro após
  falha crítica de limpeza do Excel Desktop.

### Alterado

- nomes das saídas em lote incluem a planilha de origem;
- saída, backup e log evitam colisões em execuções rápidas.

### Validado

- regressão real com três entradas aprovada no fallback: duas saídas
  válidas, uma falha estrutural isolada, original preservado e nenhum
  processo do Excel iniciado.

## 1.7.0 — 2026-07-29

### Adicionado

- distribuição portátil `onedir` para Windows x64, construída com PyInstaller,
  configuração externa editável e execução sem Python, `.venv` ou internet;
- script único de build que valida versões, conteúdo permitido, DLLs do
  pywin32, execução em caminho com espaços, diagnóstico somente leitura,
  regressão fallback, ZIP estável e checksum SHA-256;
- workflows separados para gerar candidatos de inspeção e criar uma GitHub
  Release em rascunho a partir de uma tag anotada;
- documentação própria do pacote e testes dos contratos de empacotamento.

### Alterado

- fluxo de contribuição alinhado ao uso de branches curtas para `develop` e
  branches de release para `main`;
- template de Pull Request atualizado para considerar a branch-base correta;
- arquitetura atualizada com os caminhos de versão, diagnóstico e execução
  normal, além do estado concluído da decomposição do escritor;
- documentação da governança e da regressão local alinhada ao fluxo vigente
  desde a versão 1.6.0;
- contrato do Engine alinhado ao retorno real da criação nativa, removendo uma
  segunda tentativa de recálculo que não possuía caminho executável;
- resolução das pastas adaptada ao executável congelado, mantendo configuração,
  entrada, saída, backup e logs ao lado do pacote;
- `run.ps1` e `iniciar.cmd` agora preferem o executável portátil e preservam o
  fluxo existente por código-fonte quando ele não está presente;
- diagnóstico do pywin32 ampliado para verificar `pythoncom`, `pywintypes` e
  `win32com.client` antes de considerar a integração nativa disponível.

### Validado

- 162 testes aprovados, com cobertura total de 98,44%;
- CI aprovado no Windows com Python 3.11 e Python 3.14;
- build `onedir` aprovado com PyInstaller 6.21.0 e Python 3.14.6;
- smoke do pacote aprovado em caminho com espaços e sem Python externo;
- regressões reais pelo pacote portátil aprovadas nos modos fallback e Excel
  Desktop, com 40 solicitações e 0 erros de fórmula;
- ZIP, checksum SHA-256, conteúdo permitido, pastas operacionais vazias e
  identificação do commit conferidos nos artefatos do GitHub Actions.

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
