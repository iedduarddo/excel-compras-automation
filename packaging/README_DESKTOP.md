# Excel Compras Automation — interface gráfica vX.Y.Z

Este pacote abre uma interface com botões e configurações, sem exigir Python.
Há artefatos nativos separados para Windows x64, macOS Intel e macOS Apple
Silicon. Não misture arquivos de pacotes ou arquiteturas diferentes.

## Compatibilidade inicial

- Windows 10 ou 11 x64;
- macOS 15 ou posterior em Intel x64;
- macOS 15 ou posterior em Apple Silicon arm64.

Versões anteriores podem funcionar, mas somente esses alvos são construídos e
validados automaticamente.

## Uso seguro

1. Extraia o ZIP completo em uma pasta gravável.
2. No Windows, abra `ExcelComprasAutomation.exe`. No macOS, abra
   `Excel Compras Automation.app`.
3. Use **Adicionar** para copiar planilhas `.xlsx` ou `.xlsm` para a central.
4. Escreva o pedido ou escolha **Reconhecer**, **Diagnosticar**,
   **Limpar e organizar** ou **Criar relatório**.
5. Para ações que geram arquivos, revise todo o plano exibido.
6. Clique em **Confirmar plano** somente se a prévia estiver correta. Use
   **Cancelar plano** para descartar o pedido sem criar resultado.

O aplicativo nunca sobrescreve o original. Uma confirmação gera uma nova cópia
e um backup. O monitor da interface apenas atualiza a lista de entrada; ele não
executa comandos de alteração ocultamente.

## Voz e Excel Desktop

No Windows, o botão **Voz** usa a Digitação por Voz local e devolve o texto ao
campo **Pedido** para revisão. A Tabela Dinâmica nativa depende do Microsoft
Excel Desktop.

No macOS, use o Ditado nativo do sistema diretamente no campo **Pedido**. Como
COM e `pywin32` não existem no macOS, o aplicativo produz relatórios e resumos
compatíveis por fórmulas.

## Segurança do download

Compare o SHA-256 do ZIP com o arquivo `.sha256` publicado. Estes primeiros
pacotes gráficos não possuem certificado comercial Authenticode nem Apple
Developer ID; Windows SmartScreen e macOS Gatekeeper podem exibir um aviso.
Confirme sempre a origem oficial antes de abrir.
