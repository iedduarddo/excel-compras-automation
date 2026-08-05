"""Estrutura persistente e segura usada pelo assistente local."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from src.core.exceptions import AutomationError
from src.services.files import SUPPORTED_EXTENSIONS
from src.services.text import sanitize_filename
from src.settings import PROJECT_ROOT

DEFAULT_ASSISTANT_ROOT = PROJECT_ROOT / "assistente_planilhas"


@dataclass(frozen=True, slots=True)
class AssistantConfig:
    """Preferências locais que não alteram as regras da automação."""

    candidate_name: str = ""
    use_native_pivot: bool = True
    poll_interval_seconds: float = 2.0


@dataclass(frozen=True, slots=True)
class AssistantWorkspace:
    """Caminhos do assistente, configuráveis para testes e instalações separadas."""

    root: Path = DEFAULT_ASSISTANT_ROOT

    @property
    def input_dir(self) -> Path:
        return self.root / "entrada"

    @property
    def output_dir(self) -> Path:
        return self.root / "saida"

    @property
    def backup_dir(self) -> Path:
        return self.root / "backup"

    @property
    def logs_dir(self) -> Path:
        return self.root / "logs"

    @property
    def review_dir(self) -> Path:
        return self.root / "revisao"

    @property
    def adapters_dir(self) -> Path:
        return self.root / "adaptadores"

    @property
    def commands_dir(self) -> Path:
        return self.root / "comandos"

    @property
    def pending_commands_dir(self) -> Path:
        return self.commands_dir / "pendentes"

    @property
    def completed_commands_dir(self) -> Path:
        return self.commands_dir / "concluidos"

    @property
    def failed_commands_dir(self) -> Path:
        return self.commands_dir / "falhas"

    @property
    def config_file(self) -> Path:
        return self.root / "config.json"

    def ensure(self) -> None:
        """Cria somente caminhos delimitados dentro da raiz informada."""

        for directory in (
            self.input_dir,
            self.output_dir,
            self.backup_dir,
            self.logs_dir,
            self.review_dir,
            self.adapters_dir,
            self.pending_commands_dir,
            self.completed_commands_dir,
            self.failed_commands_dir,
        ):
            directory.mkdir(parents=True, exist_ok=True)

        if not self.config_file.exists():
            self.config_file.write_text(
                json.dumps(
                    {
                        "candidate_name": "",
                        "use_native_pivot": True,
                        "poll_interval_seconds": 2,
                    },
                    ensure_ascii=False,
                    indent=2,
                )
                + "\n",
                encoding="utf-8",
            )

    def load_config(self) -> AssistantConfig:
        self.ensure()
        try:
            raw = json.loads(self.config_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError as error:
            raise AutomationError(
                f"A configuração do assistente possui JSON inválido: {error}"
            ) from error
        if not isinstance(raw, dict):
            raise AutomationError(
                "A configuração do assistente deve ser um objeto JSON."
            )

        candidate_name = raw.get("candidate_name", "")
        native = raw.get("use_native_pivot", True)
        interval = raw.get("poll_interval_seconds", 2)
        if not isinstance(candidate_name, str):
            raise AutomationError("candidate_name deve ser um texto.")
        if not isinstance(native, bool):
            raise AutomationError("use_native_pivot deve ser true ou false.")
        if not isinstance(interval, (int, float)) or not 0.5 <= float(interval) <= 60:
            raise AutomationError(
                "poll_interval_seconds deve ser um número entre 0.5 e 60."
            )
        return AssistantConfig(candidate_name.strip(), native, float(interval))

    def list_input_files(self) -> tuple[Path, ...]:
        self.ensure()
        return tuple(
            sorted(
                (
                    path.resolve()
                    for path in self.input_dir.iterdir()
                    if path.is_file()
                    and path.suffix.casefold() in SUPPORTED_EXTENSIONS
                    and not path.name.startswith("~$")
                ),
                key=lambda path: path.name.casefold(),
            )
        )

    def list_adapters(self) -> tuple[Path, ...]:
        self.ensure()
        return tuple(
            sorted(
                (path.resolve() for path in self.adapters_dir.glob("*.json")),
                key=lambda path: path.name.casefold(),
            )
        )

    def resolve_adapter(self, value: Path) -> Path:
        candidate = value.expanduser()
        if not candidate.is_absolute():
            candidate = self.adapters_dir / candidate
        candidate = candidate.resolve()
        if not candidate.is_file():
            raise AutomationError(f"Adaptador não encontrado: {candidate}")
        return candidate

    def write_json_report(
        self,
        directory: Path,
        stem: str,
        payload: dict[str, Any],
    ) -> Path:
        directory.mkdir(parents=True, exist_ok=True)
        report = directory / f"{sanitize_filename(stem)}.json"
        report.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2, default=str) + "\n",
            encoding="utf-8",
        )
        return report
