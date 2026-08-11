# Interface gráfica

A interface reutiliza o assistente e o executor universal existentes. Ela não
possui um segundo motor de automação.

## Fluxo

1. **Adicionar** copia uma ou mais planilhas `.xlsx` ou `.xlsm` para a central.
2. **Reconhecer** e **Diagnosticar** fazem somente leitura.
3. Pedidos para limpar, organizar, calcular, resumir ou criar relatório geram
   uma prévia persistente.
4. **Confirmar plano** aplica exatamente as prévias visíveis em novas cópias.
5. **Cancelar plano** encerra os planos sem gerar planilhas.

Confirmações digitadas, faladas ou colocadas na fila não são aceitas pela
interface. Essa decisão exige o botão visual e um segundo diálogo humano.

## Configurações

- nome usado pelo fluxo específico de Compras;
- uso da Tabela Dinâmica nativa no Windows;
- intervalo de atualização da lista monitorada, entre 0,5 e 60 segundos.

O monitor da interface apenas atualiza a lista quando novas planilhas aparecem.
Ele nunca executa silenciosamente um comando de alteração.

## Plataformas

- Windows 10/11 x64: interface, voz local, fallback e Excel Desktop opcional;
- macOS 15+ Intel: interface e modo compatível;
- macOS 15+ Apple Silicon: interface e modo compatível.

No macOS, o usuário pode ativar o Ditado do sistema dentro do campo **Pedido**.
A captura integrada via `Win+H` e a automação COM são exclusivas do Windows.

## Armazenamento

No desenvolvimento, a central continua em `assistente_planilhas`. No aplicativo
empacotado, os dados ficam em uma pasta gravável do usuário:

- Windows: `%LOCALAPPDATA%\ExcelComprasAutomation`;
- macOS: `~/Library/Application Support/ExcelComprasAutomation`.

Os pacotes iniciais não possuem assinatura comercial. Confira o SHA-256 antes
de abrir e mantenha todos os arquivos do ZIP da mesma arquitetura juntos.
