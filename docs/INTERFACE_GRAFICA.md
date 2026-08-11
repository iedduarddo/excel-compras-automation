# Interface gráfica

A interface reutiliza o assistente e o executor universal existentes. Ela não
possui um segundo motor de automação e mantém a confirmação humana obrigatória.

## Organização da janela

A janela possui duas abas:

- **Assistente**: concentra a importação, o pedido e a revisão do resultado;
- **Configurações**: reúne nome, Excel Desktop e intervalo do monitor sem
  congestionar a área principal.

O fluxo principal continua dividido em três etapas visíveis:

1. **Planilhas de entrada**: **Adicionar planilhas** copia arquivos `.xlsx` ou
   `.xlsm` para a central. O contador e a lista mostram o que está disponível.
2. **Descreva o pedido**: aceite linguagem natural, voz no Windows ou uma das
   ações rápidas.
3. **Revise a prévia**: confira as ações previstas e use **Confirmar plano** ou
   **Cancelar plano**.

A barra inferior informa o estado da operação e mostra progresso enquanto uma
tarefa está em execução. Nesse período, os controles que poderiam duplicar a
ação permanecem desabilitados.

## Atalhos de teclado

- `Ctrl+O` ou `Cmd+O`: adicionar planilhas;
- `F5`: atualizar a lista de entrada;
- `Ctrl+Enter` ou `Cmd+Enter`: executar o pedido escrito;
- `Ctrl+L` ou `Cmd+L`: levar o foco ao campo do pedido.

Nenhum atalho confirma alterações. A confirmação exige o botão visual e um
segundo diálogo humano.

## Fluxo seguro

1. **Adicionar planilhas** copia uma ou mais entradas para a central.
2. **Reconhecer** e **Diagnosticar** fazem somente leitura.
3. Pedidos para limpar, organizar, calcular, resumir ou criar relatório geram
   uma prévia persistente.
4. **Confirmar plano** aplica exatamente as prévias visíveis em novas cópias.
5. **Cancelar plano** encerra os planos sem gerar planilhas.

Confirmações digitadas, faladas ou colocadas na fila não são aceitas pela
interface. O monitor apenas atualiza a lista e nunca executa alterações.

## Configurações

- nome usado pelo fluxo específico de Compras;
- uso da Tabela Dinâmica nativa no Windows;
- intervalo de atualização da lista, entre 0,5 e 60 segundos.

## Plataformas

- Windows 10/11 x64: interface, voz local, fallback e Excel Desktop opcional;
- macOS 15+ Intel: interface e modo compatível;
- macOS 15+ Apple Silicon: interface e modo compatível.

No macOS, o usuário pode ativar o Ditado do sistema dentro do campo **Pedido**.
A captura integrada por voz e a automação COM são exclusivas do Windows.

## Armazenamento

No desenvolvimento, a central fica em `assistente_planilhas`. No aplicativo
empacotado, os dados ficam em uma pasta gravável do usuário:

- Windows: `%LOCALAPPDATA%\ExcelComprasAutomation`;
- macOS: `~/Library/Application Support/ExcelComprasAutomation`.

Os pacotes iniciais não possuem assinatura comercial. Confira o SHA-256 antes
de abrir e mantenha todos os arquivos do ZIP da mesma arquitetura juntos.

## Materiais Fluent e vidro

O pacote desktop prefere a interface Qt Quick. No Windows 11 22H2 ou mais
recente, a janela usa Mica e cantos arredondados fornecidos pelo DWM. Os
cartoes seguem os tokens de hierarquia, contraste e profundidade do Fluent 2.

No Windows antigo ou quando transparencia, economia de bateria ou alto
contraste desativam o efeito, a interface usa automaticamente um fundo solido
legivel. Os cartoes translucidos sao inspirados no Acrylic, mas nao sao um
backdrop Acrylic nativo. No macOS, eles mantem a mesma aparencia compativel;
essa camada nao e anunciada como Liquid Glass nativo. O efeito Apple genuino
exige uma futura casca AppKit/SwiftUI no macOS 26 ou posterior.

Use `EXCEL_COMPRAS_UI=tk` para abrir temporariamente a interface Tk de fallback.
Use `EXCEL_COMPRAS_UI=qt` para exigir Qt e falhar claramente se a dependencia
desktop nao estiver instalada.
### Alternancia de tema

O botao `Modo claro` ou `Modo escuro`, no cabecalho, troca o tema sem reiniciar
a automacao. A preferencia e salva pelo Qt para a proxima abertura. O fundo
mantem uma camada uniforme sobre o Mica para evitar falhas de contraste ou
faixas sem composicao durante redimensionamentos no Windows.
