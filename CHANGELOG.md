# Histórico de versões

Todas as mudanças relevantes deste projeto serão registradas neste arquivo.
O formato segue a ideia do [Keep a Changelog](https://keepachangelog.com/),
com versionamento semântico.

## Em desenvolvimento

### Planejado

- separar o escritor da planilha em componentes menores;
- avaliar um runner Windows próprio para testes de integração com Excel Desktop;
- preparar empacotamento e interface para usuários não técnicos.

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
