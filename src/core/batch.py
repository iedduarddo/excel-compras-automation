"""Orquestra o processamento sequencial de várias planilhas."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from src.core.engine import AutomationEngine
from src.core.exceptions import AutomationError, ExcelDesktopCleanupError
from src.core.models import RunResult
from src.services.files import list_input_files, validate_candidate_name


@dataclass(frozen=True)
class BatchItemResult:
    """Resultado de uma entrada individual do lote."""

    input_file: Path
    result: RunResult | None = None
    error: str | None = None

    @property
    def succeeded(self) -> bool:
        return self.result is not None


@dataclass(frozen=True)
class BatchResult:
    """Resumo agregado de um processamento em lote."""

    items: tuple[BatchItemResult, ...]

    @property
    def succeeded(self) -> int:
        return sum(item.succeeded for item in self.items)

    @property
    def failed(self) -> int:
        return len(self.items) - self.succeeded


class BatchAutomation:
    """Executa todas as entradas válidas e isola falhas não críticas."""

    def __init__(self, engine: AutomationEngine | None = None) -> None:
        self.engine = engine or AutomationEngine()

    def run(
        self,
        *,
        candidate_name: str,
        use_native_pivot: bool = True,
        verbose: bool = False,
        adapter: str | Path | None = None,
    ) -> BatchResult:
        """Processa as planilhas em ordem alfabética e retorna o resumo."""

        candidate_name = validate_candidate_name(candidate_name)
        items: list[BatchItemResult] = []
        for input_file in list_input_files():
            try:
                run_options: dict[str, object] = {
                    "input_value": input_file,
                    "candidate_name": candidate_name,
                    "use_native_pivot": use_native_pivot,
                    "verbose": verbose,
                    "output_label": input_file.stem,
                }
                if adapter is not None:
                    run_options["adapter"] = adapter
                result = self.engine.run(**run_options)
            except ExcelDesktopCleanupError:
                raise
            except AutomationError as error:
                items.append(
                    BatchItemResult(
                        input_file=input_file,
                        error=str(error),
                    )
                )
            else:
                items.append(
                    BatchItemResult(
                        input_file=input_file,
                        result=result,
                    )
                )

        return BatchResult(tuple(items))
