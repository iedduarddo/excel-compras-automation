# Como contribuir

Este guia descreve o fluxo de desenvolvimento do Excel Compras Automation.
Ele também serve como roteiro para manter o projeto sozinho, sem enviar
alterações diretamente para a versão estável.

## Princípios do projeto

- A branch `main` representa a versão estável publicada.
- A branch `develop` reúne mudanças já revisadas para a próxima versão.
- Funcionalidades, correções, testes e documentação são desenvolvidos em
  branches curtas criadas a partir da `develop`.
- Toda branch curta chega à `develop` por Pull Request.
- Toda publicação usa uma branch `release/vX.Y.Z`, integrada à `main` por Pull
  Request.
- Não faça commits diretamente em `main` ou `develop`. A única atualização
  direta permitida em `develop` é o fast-forward final da `main` depois de uma
  release validada.
- Os checks `Python 3.11` e `Python 3.14` devem ser aprovados.
- O merge deve ficar preso ao SHA revisado com `--match-head-commit`.
- Branches temporárias só são removidas depois da aprovação do CI pós-merge.
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

## 2. Sincronizar a base e criar uma branch curta

Comece sempre com a árvore de trabalho limpa:

```powershell
if (git status --porcelain) {
    throw "Existem alterações locais. Operação interrompida."
}
```

Atualize a `develop` sem criar merge local:

```powershell
git fetch origin --prune
git switch develop
git pull --ff-only origin develop
```

Confirme a sincronização e crie uma branch com um único objetivo:

```powershell
if (
    (git rev-parse HEAD).Trim() -ne
    (git rev-parse origin/develop).Trim()
) {
    throw "A develop local e a remota não estão sincronizadas."
}

git switch -c feat/nome-da-funcionalidade
```

Use o prefixo correspondente ao objetivo: `feat/`, `fix/`, `refactor/`,
`test/`, `docs/`, `ci/` ou `chore/`.

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

## 5. Criar o commit e publicar a branch curta

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

Depois, publique a branch atual sem alterar diretamente `develop` ou `main`:

```powershell
$branch = (git branch --show-current).Trim()

if ($branch -in @("main", "develop")) {
    throw "Não faça commits de desenvolvimento diretamente em $branch."
}

git commit -m "tipo: descrição objetiva"
git push -u origin $branch
```

Confirme que o commit local é o mesmo que foi publicado:

```powershell
git fetch origin --prune

if (
    (git rev-parse HEAD).Trim() -ne
    (git rev-parse "origin/$branch").Trim()
) {
    throw "A branch local e a remota não estão sincronizadas."
}

git status -sb
```

## 6. Integrar uma branch curta à `develop`

Crie o PR inicialmente como rascunho:

```powershell
$branch = (git branch --show-current).Trim()

gh pr create `
    --draft `
    --base develop `
    --head $branch `
    --title "Título objetivo" `
    --body-file CAMINHO_DO_CORPO.md
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
5. confirme que o PR está sem conflitos;
6. marque o PR como pronto.

O projeto individual não exige aprovação de outra pessoa, mas exige
autorrevisão e CI aprovado.

O merge deve ser preso ao commit efetivamente revisado:

```powershell
$pr = NUMERO_DO_PR
$headDoPr = (
    gh pr view $pr --json headRefOid --jq ".headRefOid"
).Trim()

gh pr ready $pr
gh pr merge $pr `
    --merge `
    --match-head-commit $headDoPr
```

Se a `develop` mudar durante a revisão, atualize a branch e execute novamente
todos os checks antes do merge.

## 7. Confirmar o CI pós-merge e remover a branch curta

Obtenha o SHA do merge e aguarde a execução de `push` da workflow `ci.yml` na
`develop` para esse mesmo commit:

```powershell
$mergeSha = (
    gh pr view $pr --json mergeCommit --jq ".mergeCommit.oid"
).Trim()

gh run watch ID_DA_EXECUCAO --exit-status
```

Somente depois do CI aprovado, atualize a `develop`, confirme a integração e
remova a branch temporária:

```powershell
git fetch origin --prune
git switch develop
git pull --ff-only origin develop

if ((git rev-parse HEAD).Trim() -ne $mergeSha) {
    throw "A develop não terminou no merge validado."
}

git merge-base --is-ancestor $branch develop

if ($LASTEXITCODE -ne 0) {
    throw "A branch curta não está completamente integrada."
}

git push origin --delete $branch
git branch -d -- $branch

git status -sb
```

Preserve a branch quando o CI pós-merge falhar ou quando a integração não puder
ser comprovada.

## 8. Preparar uma versão

Crie a branch de release somente depois que todas as mudanças planejadas
estiverem integradas e o CI da `develop` estiver aprovado:

```powershell
git fetch origin --prune
git switch develop
git pull --ff-only origin develop
git switch -c release/vX.Y.Z
```

Atualize:

- `pyproject.toml`;
- `src/__init__.py`;
- `CHANGELOG.md`;
- `README.md`, quando houver instruções ou números específicos da versão.

Quando a mudança afeta a distribuição Windows, valide também:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv\Scripts\python.exe scripts\build_portable.py `
    --expected-version X.Y.Z `
    --commit (git rev-parse HEAD)
```

Esse comando executa o build e o smoke test sem Excel, depois cria em `dist` o
ZIP `ExcelComprasAutomation-vX.Y.Z-windows-x64.zip` e o respectivo `.sha256`.
Esses arquivos são gerados, não rastreados pelo Git.

Abra um PR da branch `release/vX.Y.Z` para `main`. Crie a tag anotada somente
depois do merge, do CI de `push` da `main` e sobre o commit final validado:

```powershell
git switch main
git pull --ff-only origin main
git tag -a vX.Y.Z COMMIT_FINAL -m "Versão X.Y.Z"
git push origin vX.Y.Z
```

Ao receber a tag, `release-windows.yml` confirma que ela é anotada e aponta para
a `main`, reconstrói o pacote com Python 3.14.6 e cria uma GitHub Release em
rascunho com o ZIP e o checksum. Baixe exatamente esse asset, confira o SHA-256
e execute as regressões fallback e nativa em uma máquina Windows x64. Publique o
rascunho somente depois dessa validação, sem substituir os assets.

Depois, sincronize `develop` com `main` por fast-forward, aguarde o CI final da
`develop` e somente então remova a branch de release. Ao final, `main`,
`develop`, `origin/main` e `origin/develop` devem apontar para o mesmo commit.

Nunca inclua planilhas de entrada, saídas, backups ou logs nos anexos da
release. Para versões com distribuição portátil, os únicos assets esperados são
o ZIP Windows x64 e seu `.sha256`.

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
