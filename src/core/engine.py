"""Orquestrador principal do fluxo de automação."""

from __future__ import annotations

from pathlib import Path

from src.business.policies import analyze_travels, find_last_data_row, read_policies
from src.business.priorities import apply_priority_scores
from src.core.exceptions import ExcelDesktopError
from src.core.models import RunResult
from src.excel.detection import WorkbookDetector
from src.excel.excel_desktop import (
    create_native_pivot_and_recalculate,
    pywin32_is_available,
    recalculate_only,
)
from src.excel.validation import validate_output
from src.excel.workbook_writer import (
    apply_conditional_formatting,
    create_fallback_summary_and_chart,
    create_support_sheet,
    ensure_derived_columns,
    finalize_workbook,
    load_source_workbook,
    write_base_formulas,
    write_responses,
)
from src.services.files import (
    prepare_run_paths,
    resolve_input_file,
    validate_candidate_name,
)
from src.services.logging_setup import configure_logging
from src.settings import load_aliases, load_rules


class AutomationEngine:
    """Executa todas as etapas sem concentrar as regras no arquivo main."""

    def run(
        self,
        input_value: str | Path | None,
        candidate_name: str,
        use_native_pivot: bool = True,
        verbose: bool = False,
    ) -> RunResult:
        """Processa uma cópia, recalcula no Excel e valida a entrega."""

        input_file = resolve_input_file(input_value)
        candidate_name = validate_candidate_name(candidate_name)
        paths = prepare_run_paths(input_file, candidate_name)
        logger = configure_logging(paths.log_file, verbose=verbose)
        logger.info("Iniciando Excel Compras Automation.")
        logger.info("Entrada: %s", paths.input_file)
        logger.info("Backup criado: %s", paths.backup_file)

        aliases = load_aliases()
        rules = load_rules()
        workbook = load_source_workbook(paths.input_file)
        try:
            logger.info("Identificando abas e cabeçalhos por conteúdo.")
            layout = WorkbookDetector(aliases).detect(workbook)
            logger.info(
                "Abas identificadas | base=%s | políticas=%s | respostas=%s",
                layout.base.title,
                layout.policies.title,
                layout.responses.title,
            )
            logger.debug("Colunas da base: %s", layout.base.columns)
            logger.debug("Colunas de políticas: %s", layout.policies.columns)
            logger.debug("Colunas de respostas: %s", layout.responses.columns)

            policies = read_policies(workbook, layout)
            last_row = find_last_data_row(workbook, layout)
            travels = analyze_travels(workbook, layout, policies)
            apply_priority_scores(travels, rules)
            logger.info(
                "%d solicitações e %d políticas validadas.",
                len(travels),
                len(policies),
            )

            layout = ensure_derived_columns(workbook, layout, last_row)
            support_refs = create_support_sheet(
                workbook,
                layout,
                travels,
                rules,
                last_row,
            )
            write_base_formulas(
                workbook,
                layout,
                policies,
                support_refs,
                last_row,
            )
            pivot_start_row = write_responses(
                workbook,
                layout,
                aliases,
                travels,
                support_refs,
                last_row,
                top_quantity=int(rules["top_requests"]),
            )
            apply_conditional_formatting(workbook, layout, last_row)
            finalize_workbook(workbook)
            workbook.save(paths.output_file)
        finally:
            workbook.close()
        logger.info("Primeira versão gravada: %s", paths.output_file)

        native_pivot_created = False
        desktop_recalculated = False
        native_attempt_failed = False
        native_available = use_native_pivot and pywin32_is_available()
        if native_available:
            try:
                desktop_recalculated = create_native_pivot_and_recalculate(
                    paths.output_file,
                    layout,
                    pivot_start_row,
                    last_row,
                    logger,
                )
                native_pivot_created = True
            except ExcelDesktopError as error:
                native_attempt_failed = True
                logger.warning("%s", error)
                logger.warning(
                    "Será criado o resumo formula-driven de compatibilidade."
                )

        if not native_pivot_created:
            fallback_workbook = load_source_workbook(paths.output_file)
            try:
                create_fallback_summary_and_chart(
                    fallback_workbook,
                    layout,
                    travels,
                    pivot_start_row,
                    last_row,
                )
                finalize_workbook(fallback_workbook)
                fallback_workbook.save(paths.output_file)
            finally:
                fallback_workbook.close()
            logger.info("Resumo e gráfico de compatibilidade criados.")
            if native_available and not native_attempt_failed:
                try:
                    recalculate_only(paths.output_file, logger)
                    desktop_recalculated = True
                except Exception:
                    logger.warning(
                        "Não foi possível recalcular o resumo pelo Excel Desktop. "
                        "O Excel recalculará as fórmulas quando o arquivo for aberto.",
                        exc_info=True,
                    )

        logger.info("Executando validações finais.")
        checks = validate_output(
            paths.output_file,
            layout,
            aliases,
            travels,
            last_row,
            native_pivot_expected=native_pivot_created,
            cached_values_expected=desktop_recalculated,
        )
        logger.info("Validações concluídas: %s", checks)
        logger.info("Automação concluída com sucesso.")

        return RunResult(
            output_file=paths.output_file,
            backup_file=paths.backup_file,
            log_file=paths.log_file,
            native_pivot_created=native_pivot_created,
            detected_sheets={
                "base": layout.base.title,
                "policies": layout.policies.title,
                "responses": layout.responses.title,
            },
            detected_columns={
                "base": dict(layout.base.columns),
                "policies": dict(layout.policies.columns),
                "responses": dict(layout.responses.columns),
            },
            checks=checks,
        )
