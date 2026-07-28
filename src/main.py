"""Interface de linha de comando do projeto."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from src import __version__
from src.core.engine import AutomationEngine
from src.core.exceptions import AutomationError
from src.services.diagnostics import format_diagnostic_report, run_diagnostics


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=("Preenche e analisa a planilha do teste de Analista de Compras.")
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help=(
            "Caminho da planilha. Se omitido, será usada a única planilha "
            "existente na pasta input."
        ),
    )
    parser.add_argument(
        "--candidate-name",
        "--nome",
        dest="candidate_name",
        default=None,
        help="Nome completo que será usado no arquivo de saída.",
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
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        if args.diagnostic:
            report = run_diagnostics(args.input)
            print(format_diagnostic_report(report))
            return report.exit_code

        candidate_name = args.candidate_name
        if not candidate_name:
            candidate_name = input("Digite seu nome completo: ").strip()

        result = AutomationEngine().run(
            input_value=args.input,
            candidate_name=candidate_name,
            use_native_pivot=not args.sem_pivot_nativo,
            verbose=args.verbose,
        )
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


if __name__ == "__main__":
    raise SystemExit(main())
