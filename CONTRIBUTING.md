# Como contribuir

Este guia descreve o fluxo de desenvolvimento do Excel Compras Automation.
Ele também serve como roteiro para manter o projeto sozinho, sem enviar
alterações diretamente para a versão estável.

## Princípios do projeto

- A branch `main` representa a versão estável.
- A branch `develop` recebe as mudanças antes da publicação.
- Toda mudança na `main` deve passar por Pull Request.
- Os checks `Python 3.11` e `Python 3.14` devem ser aprovados.
- Arquivos reais das pastas `input`, `output`, `backup` e `logs` não devem ser
  enviados ao GitHub.
- Alterações no Excel devem preservar os dados originais e as fórmulas exigidas.

## 1. Preparar o ambiente

Na primeira utilização:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

Nas próximas utilizações, confirme o ambiente:

```powershell
.\.venv\Scripts\python.exe --version
.\.venv\Scripts\python.exe -m pip check
```

## 2. Sincronizar as branches

Comece sempre com a árvore de trabalho limpa:

```powershell
git status --short
```

Depois sincronize:

```powershell
git fetch origin

git switch main
git pull --ff-only origin main

git switch develop
git merge --ff-only main
git push origin develop
```

Não faça commits diretamente na `main`. A proteção do GitHub rejeita mudanças
que não tenham passado pelo Pull Request e pelos checks obrigatórios.

## 3. Implementar uma mudança

Faça somente alterações relacionadas ao mesmo objetivo. Antes do commit,
confira:

```powershell
git status --short
git diff --check
git diff
```

Nunca inclua:

- planilhas reais ou resultados gerados;
- arquivos de backup e log;
- `.venv`;
- credenciais, tokens ou e-mails particulares;
- arquivos temporários do Excel.

## 4. Executar as validações

Use o interpretador da `.venv` para não depender da ativação do terminal:

```powershell
.\.venv\Scripts\python.exe -m compileall -q main.py src
.\.venv\Scripts\python.exe -m ruff check .
.\.venv\Scripts\python.exe -m ruff format --check .
.\.venv\Scripts\python.exe -m pytest --cov=src --cov-report=term-missing -q
```

Se a mudança afetar leitura, escrita, fórmulas, formatação, PivotTable ou
gráficos, execute também a regressão com Microsoft Excel Desktop:

```powershell
powershell -ExecutionPolicy Bypass -File .\run.ps1 `
    -NomeCompleto "NOME COMPLETO PARA TESTE"
```

Confira o arquivo gerado e confirme no log:

- número correto de solicitações;
- zero erros de fórmula;
- Tabela Dinâmica e gráfico presentes;
- original preservado.

## 5. Criar o commit

Adicione somente os arquivos da mudança:

```powershell
git add CAMINHO_DO_ARQUIVO
git diff --cached --check
git diff --cached --stat
```

Use mensagens curtas e objetivas:

```text
feat: adiciona uma funcionalidade
fix: corrige um comportamento
test: amplia os testes
docs: atualiza a documentação
refactor: melhora a estrutura sem mudar o resultado
ci: altera a integração contínua
chore: executa manutenção do projeto
```

Depois:

```powershell
git commit -m "tipo: descrição objetiva"
git push origin develop
```

## 6. Abrir o Pull Request

Crie o PR de `develop` para `main`, inicialmente como rascunho:

```powershell
gh pr create --draft --base main --head develop
```

O modelo do repositório apresentará os campos que devem ser preenchidos.
Acompanhe o CI:

```powershell
gh pr checks --watch --interval 10
```

Antes do merge:

1. confira a aba **Files changed**;
2. confirme que não há arquivos fora do escopo;
3. confirme os checks `Python 3.11` e `Python 3.14`;
4. resolva todas as conversas;
5. marque o PR como pronto.

O projeto individual não exige aprovação de outra pessoa, mas exige
autorrevisão e CI aprovado.

## 7. Finalizar e sincronizar

Depois do merge:

```powershell
git fetch origin

git switch main
git pull --ff-only origin main

git switch develop
git merge --ff-only main
git push origin develop

git status -sb
```

Ao final, `main`, `develop`, `origin/main` e `origin/develop` devem apontar para
o mesmo commit.

## 8. Publicar uma versão

Crie uma branch de release a partir da `develop` sincronizada:

```powershell
git switch develop
git switch -c release/vX.Y.Z
```

Atualize:

- `pyproject.toml`;
- `src/__init__.py`;
- `CHANGELOG.md`.

Abra um PR da branch `release/vX.Y.Z` para `main`. Crie a tag anotada somente
depois do merge e sobre o commit final da `main`:

```powershell
git switch main
git pull --ff-only origin main
git tag -a vX.Y.Z -m "Versão X.Y.Z"
git push origin vX.Y.Z
```

Nunca inclua a planilha de entrada ou o arquivo de saída nos anexos da release.
O código-fonte e a documentação do repositório são suficientes.

## Relatar um problema

Ao abrir uma issue, informe:

- comando executado;
- mensagem completa do erro;
- versão do Python;
- versão do Microsoft Excel, quando aplicável;
- se o arquivo estava aberto;
- passos mínimos para reproduzir.

Remova dados pessoais, nomes de empresas, valores confidenciais e conteúdo real
da planilha antes de publicar qualquer evidência.
