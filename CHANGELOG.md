# Histórico de versões

Todas as mudanças relevantes deste projeto serão registradas neste arquivo.
O formato segue a ideia do [Keep a Changelog](https://keepachangelog.com/),
com versionamento semântico.

## Em desenvolvimento

### Adicionado

- Ruff para análise estática, ordenação de imports e formatação;
- pytest-cov para medição da cobertura dos testes;
- configuração reproduzível das ferramentas no `pyproject.toml`.

### Corrigido

- imports não utilizados e fora da ordem;
- tratamento silencioso de erro durante a restauração do Excel;
- criação e fechamento do arquivo temporário de reparo OOXML;
- geração de timestamps com informação de fuso horário.

### Planejado

- ampliar os testes de regressão e integração;
- separar o escritor da planilha em componentes menores;
- fortalecer a validação de configurações e artefatos do Excel;
- adicionar análise estática, cobertura e integração contínua;
- preparar empacotamento e interface para usuários não técnicos.

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
