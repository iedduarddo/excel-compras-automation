"""Interface de linha de comando do projeto."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from src import __version__
from src.assistant.service import FolderAssistant, format_assistant_result
from src.assistant.voice import recognize_voice
from src.assistant.workspace import AssistantWorkspace
from src.core.batch import BatchAutomation, BatchResult
from src.core.engine import AutomationEngine
from src.core.exceptions import AutomationError
from src.services.diagnostics import format_diagnostic_report, run_diagnostics


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=("Preenche e analisa a planilha do teste de Analista de Compras.")
    )
    input_mode = parser.add_mutually_exclusive_group()
    input_mode.add_argument(
        "--input",
        type=Path,
        default=None,
        help=(
            "Caminho da planilha. Se omitido, será usada a única planilha "
            "existente na pasta input."
        ),
    )
    input_mode.add_argument(
        "--lote",
        "--batch",
        dest="batch",
        action="store_true",
        help="Processa todas as planilhas válidas da pasta input.",
    )
    parser.add_argument(
        "--candidate-name",
        "--nome",
        dest="candidate_name",
        default=None,
        help="Nome completo que será usado no arquivo de saída.",
    )
    parser.add_argument(
        "--adaptador",
        "--adapter",
        dest="adapter",
        type=Path,
        default=None,
        help=(
            "Caminho de um perfil JSON que acrescenta nomes de abas, "
            "colunas e indicadores ao mapeamento padrão."
        ),
    )
    parser.add_argument(
        "--sem-pivot-nativo",
        action="store_true",
        help=(
            "Não usa o Excel Desktop. Cria um resumo formula-driven compatível "
            "com outros ambientes."
        ),
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Mostra informações técnicas adicionais no terminal.",
    )
    parser.add_argument(
        "--diagnostico",
        "--diagnostic",
        dest="diagnostic",
        action="store_true",
        help=(
            "Verifica o ambiente e a estrutura da entrada sem criar backup ou saída."
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    parser.add_argument(
        "--assistente",
        "--assistant",
        dest="assistant",
        action="store_true",
        help="Usa a pasta monitorada e a fila de comandos escritos.",
    )
    parser.add_argument(
        "--comando",
        "--command",
        dest="command",
        nargs="+",
        default=None,
        help="Executa diretamente um comando conhecido do assistente.",
    )
    parser.add_argument(
        "--monitorar",
        "--watch",
        dest="watch",
        action="store_true",
        help="Mantém o assistente monitorando entradas e comandos.",
    )
    parser.add_argument(
        "--voz",
        "--voice",
        dest="voice",
        action="store_true",
        help="Escuta um comando pelo reconhecedor de fala local do Windows.",
    )
    parser.add_argument(
        "--preparar-pastas",
        dest="prepare_assistant",
        action="store_true",
        help="Cria a estrutura do assistente sem processar arquivos.",
    )
    parser.add_argument(
        "--pasta-assistente",
        dest="assistant_root",
        type=Path,
        default=None,
        help="Raiz alternativa para as pastas monitoradas.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.batch and args.diagnostic:
        parser.error("Use apenas um modo por vez: --lote ou --diagnostico.")
    assistant_options = (
        args.command
        or args.watch
        or args.voice
        or args.prepare_assistant
        or args.assistant_root
    )
    if assistant_options and not args.assistant:
        parser.error("Use --assistente junto das opções de comando ou monitoramento.")
    assistant_actions = sum(
        bool(option)
        for option in (args.command, args.watch, args.voice, args.prepare_assistant)
    )
    if assistant_actions > 1:
        parser.error("Use apenas uma ação do assistente por execução.")
    if args.assistant and (
        args.input
        or args.batch
        or args.diagnostic
        or args.candidate_name
        or args.adapter
        or args.sem_pivot_nativo
        or args.verbose
    ):
        parser.error(
            "O assistente recebe opções pelo comando ou config.json; não combine "
            "com os parâmetros do modo tradicional."
        )

    try:
        if args.assistant:
            workspace = (
                AssistantWorkspace(args.assistant_root)
                if args.assistant_root is not None
                else AssistantWorkspace()
            )
            assistant = FolderAssistant(workspace)
            assistant.initialize()
            if args.prepare_assistant:
                print(f"Pastas do assistente preparadas em: {workspace.root}")
                return 0
            if args.command:
                result = assistant.execute(" ".join(args.command))
                print(format_assistant_result(result))
                return 0 if result.succeeded else 1
            if args.voice:
                print("Escutando um comando... fale agora.")
                recognition = recognize_voice()
                details = f"idioma {recognition.culture}"
                if recognition.confidence is not None:
                    details += f", confiança {recognition.confidence:.0%}"
                print(f'Comando confirmado por voz: "{recognition.text}" ({details})')
                result = assistant.execute(recognition.text)
                print(format_assistant_result(result))
                return 0 if result.succeeded else 1
            if args.watch:
                print(f"Monitorando: {workspace.root}")
                assistant.watch()
                return 0

            results = assistant.run_pending_once()
            if not results:
                print(
                    "Nenhum comando pendente. Adicione um arquivo .txt em "
                    f"{workspace.pending_commands_dir}."
                )
                return 0
            for result in results:
                print(format_assistant_result(result))
            return 0 if all(result.succeeded for result in results) else 1

        if args.diagnostic:
            if args.adapter is None:
                report = run_diagnostics(args.input)
            else:
                report = run_diagnostics(args.input, adapter=args.adapter)
            print(format_diagnostic_report(report))
            return report.exit_code

        candidate_name = args.candidate_name
        if not candidate_name:
            candidate_name = input("Digite seu nome completo: ").strip()

        if args.batch:
            batch_options: dict[str, object] = {
                "candidate_name": candidate_name,
                "use_native_pivot": not args.sem_pivot_nativo,
                "verbose": args.verbose,
            }
            if args.adapter is not None:
                batch_options["adapter"] = args.adapter
            batch_result = BatchAutomation().run(**batch_options)
            _print_batch_result(batch_result)
            return 1 if batch_result.failed else 0

        run_options: dict[str, object] = {
            "input_value": args.input,
            "candidate_name": candidate_name,
            "use_native_pivot": not args.sem_pivot_nativo,
            "verbose": args.verbose,
        }
        if args.adapter is not None:
            run_options["adapter"] = args.adapter
        result = AutomationEngine().run(**run_options)
    except AutomationError as error:
        print("\nERRO: a automação não foi concluída.", file=sys.stderr)
        print(str(error), file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nExecução cancelada pelo usuário.", file=sys.stderr)
        return 130

    print("\n" + "=" * 72)
    print("AUTOMAÇÃO CONCLUÍDA COM SUCESSO")
    print("=" * 72)
    print(f"Arquivo final : {result.output_file}")
    print(f"Backup        : {result.backup_file}")
    print(f"Log           : {result.log_file}")
    print(
        "Tabela dinâmica: "
        + (
            "nativa do Excel"
            if result.native_pivot_created
            else "resumo compatível com fórmulas"
        )
    )
    print(
        f"Validação     : {result.checks['travel_rows']} solicitações, "
        f"{result.checks['formula_errors']} erros de fórmula"
    )
    print("=" * 72)
    return 0


def _print_batch_result(batch_result: BatchResult) -> None:
    """Exibe um resumo único, mantendo os detalhes de cada entrada."""

    print("\n" + "=" * 72)
    print("PROCESSAMENTO EM LOTE CONCLUÍDO")
    print("=" * 72)
    print(f"Planilhas     : {len(batch_result.items)}")
    print(f"Sucessos      : {batch_result.succeeded}")
    print(f"Falhas        : {batch_result.failed}")
    print("-" * 72)
    for item in batch_result.items:
        if item.result is not None:
            print(f"[OK] {item.input_file.name}")
            print(f"     Saída: {item.result.output_file}")
        else:
            print(f"[FALHA] {item.input_file.name}")
            print(f"        Motivo: {item.error}")
    print("=" * 72)


if __name__ == "__main__":
    raise SystemExit(main())
