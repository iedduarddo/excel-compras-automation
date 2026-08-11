"""Janela desktop multiplataforma do Excel Compras Automation."""

from __future__ import annotations

import os
import queue
import sys
from collections.abc import Callable
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import TypeVar

from src import __version__
from src.assistant.service import AssistantResult, format_assistant_result
from src.core.exceptions import AutomationError
from src.gui.controller import DesktopController

_T = TypeVar("_T")


def calculate_window_geometry(screen_width: int, screen_height: int) -> str:
    """Calcula uma janela centralizada que tambem cabe em telas menores."""

    available_width = max(760, screen_width - 64)
    available_height = max(560, screen_height - 112)
    width = min(1180, available_width)
    height = min(780, available_height)
    left = max((screen_width - width) // 2, 0)
    top = max((screen_height - height) // 2, 0)
    return f"{width}x{height}+{left}+{top}"


def platform_description(platform_name: str) -> str:
    """Retorna um nome curto e compreensivel para a plataforma atual."""

    if platform_name == "win32":
        return "Windows"
    if platform_name == "darwin":
        return "macOS"
    return "Desktop"


def run_desktop_app(root: Path | None = None) -> int:
    """Prefere Qt Quick e conserva Tk como fallback explicito."""

    mode = os.environ.get("EXCEL_COMPRAS_UI", "auto").strip().casefold()
    if mode == "tk":
        return run_tk_desktop_app(root)
    if mode not in {"auto", "qt"}:
        print("EXCEL_COMPRAS_UI deve ser auto, qt ou tk.", file=sys.stderr)
        return 2
    try:
        from src.gui.qt_app import run_qt_desktop_app

        return run_qt_desktop_app(root)
    except ModuleNotFoundError as error:
        missing = error.name and error.name.split(".", maxsplit=1)[0] == "PySide6"
        if not missing:
            raise
        if mode == "qt":
            print(
                "PySide6 nao esta instalado. Use requirements-desktop-build.txt.",
                file=sys.stderr,
            )
            return 1
        return run_tk_desktop_app(root)


def run_tk_desktop_app(root: Path | None = None) -> int:
    """Abre a janela somente quando o modo grafico e solicitado."""

    import tkinter as tk
    from tkinter import messagebox

    try:
        root_window = tk.Tk()
    except tk.TclError as error:
        print(f"Nao foi possivel abrir a interface grafica: {error}", file=sys.stderr)
        return 1
    try:
        controller = DesktopController(root)
        app = DesktopApp(root_window, controller)
        app.root.mainloop()
        return 0
    except AutomationError as error:
        messagebox.showerror("Excel Compras Automation", str(error), parent=root_window)
        root_window.destroy()
        return 1


class DesktopApp:
    """Interface orientada a tarefas, com execucao em segundo plano."""

    def __init__(self, root: object, controller: DesktopController) -> None:
        import tkinter as tk
        from tkinter import ttk

        self.tk = tk
        self.ttk = ttk
        self.root = root
        self.controller = controller
        self.executor = ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="excel-gui"
        )
        self.completed: queue.SimpleQueue[
            tuple[Future[object], Callable[[object], None] | None]
        ] = queue.SimpleQueue()
        self.current_plan_ids: tuple[str, ...] = ()
        self.monitoring = False
        self.monitor_after_id: str | None = None
        self.busy = False
        self.general_buttons: list[object] = []

        self.root.title(f"Excel Compras Automation {__version__}")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(calculate_window_geometry(screen_width, screen_height))
        self.root.minsize(min(880, screen_width - 32), min(620, screen_height - 72))
        self.root.protocol("WM_DELETE_WINDOW", self._close)

        self.status = tk.StringVar(value="Pronto para começar")
        self.input_summary = tk.StringVar(value="Nenhuma planilha adicionada")
        self.plan_summary = tk.StringVar(value="Nenhum plano aguardando")
        self.command = tk.StringVar(value="reconhecer todas")
        self.candidate_name = tk.StringVar()
        self.native_pivot = tk.BooleanVar(value=True)
        self.poll_interval = tk.StringVar(value="2")

        self._configure_styles()
        self._build()
        self._load_settings()
        self.refresh_inputs()
        self.root.after(100, self._poll_completed)
        self.root.after(150, self.command_entry.focus_set)

    def _configure_styles(self) -> None:
        from tkinter import font, ttk

        style = ttk.Style(self.root)
        default_font = font.nametofont("TkDefaultFont").copy()
        default_font.configure(size=10)
        heading_font = font.nametofont("TkHeadingFont").copy()
        heading_font.configure(size=11, weight="bold")
        title_font = font.nametofont("TkHeadingFont").copy()
        title_font.configure(size=20, weight="bold")
        subtitle_font = font.nametofont("TkDefaultFont").copy()
        subtitle_font.configure(size=10)
        action_font = font.nametofont("TkDefaultFont").copy()
        action_font.configure(weight="bold")

        style.configure(".", font=default_font)
        style.configure("Title.TLabel", font=title_font)
        style.configure("Subtitle.TLabel", font=subtitle_font)
        style.configure("Section.TLabel", font=heading_font)
        style.configure("Step.TLabel", font=heading_font, padding=(8, 3))
        style.configure("Card.TLabelframe", padding=14)
        style.configure("Card.TLabelframe.Label", font=heading_font)
        style.configure("Primary.TButton", font=action_font, padding=(16, 9))
        style.configure("Action.TButton", padding=(12, 8))
        style.configure("Quiet.TButton", padding=(10, 7))
        style.configure("Plan.TButton", font=action_font, padding=(14, 8))
        style.configure("Modern.TEntry", padding=7)
        style.configure("Modern.TNotebook", tabmargins=(0, 8, 0, 0))
        style.configure("Modern.TNotebook.Tab", padding=(18, 9))
        style.configure("Status.TLabel", padding=(2, 5))
        style.configure("StatusSuccess.TLabel", padding=(2, 5))
        style.configure("StatusWarning.TLabel", padding=(2, 5))
        style.configure("StatusError.TLabel", padding=(2, 5))

    def _build(self) -> None:
        from tkinter import ttk

        shell = ttk.Frame(self.root, padding=(20, 16, 20, 14))
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        self._build_header(shell)

        self.notebook = ttk.Notebook(shell, style="Modern.TNotebook")
        self.notebook.grid(row=1, column=0, sticky="nsew")
        assistant_tab = ttk.Frame(self.notebook, padding=(0, 14, 0, 0))
        settings_tab = ttk.Frame(self.notebook, padding=18)
        self.notebook.add(assistant_tab, text="Assistente")
        self.notebook.add(settings_tab, text="Configurações")

        self._build_assistant_tab(assistant_tab)
        self._build_settings_tab(settings_tab)
        self._build_status_bar(shell)
        self._bind_shortcuts()

    def _build_header(self, parent: object) -> None:
        from tkinter import ttk

        header = ttk.Frame(parent)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(
            header,
            text="Excel Compras Automation",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Importe, descreva e confirme antes de gerar novas cópias.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))

        environment = ttk.Frame(header)
        environment.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Label(
            environment,
            text=f"v{__version__}",
            style="Section.TLabel",
        ).grid(row=0, column=0, sticky="e")
        ttk.Label(
            environment,
            text=(
                f"{platform_description(self.controller.capabilities.platform)} · "
                "modo seguro"
            ),
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="e", pady=(3, 0))

    def _build_assistant_tab(self, parent: object) -> None:
        from tkinter import ttk

        parent.columnconfigure(0, weight=2)
        parent.columnconfigure(1, weight=3)
        parent.rowconfigure(0, weight=1)

        self._build_inputs_panel(parent)
        workspace = ttk.Frame(parent)
        workspace.grid(row=0, column=1, sticky="nsew", padx=(14, 0))
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(1, weight=1)
        self._build_request_panel(workspace)
        self._build_preview_panel(workspace)

    def _build_inputs_panel(self, parent: object) -> None:
        from tkinter import ttk

        files = ttk.LabelFrame(
            parent,
            text="1. Planilhas de entrada",
            style="Card.TLabelframe",
        )
        files.grid(row=0, column=0, sticky="nsew")
        files.columnconfigure(0, weight=1)
        files.rowconfigure(2, weight=1)

        ttk.Label(
            files,
            text="Adicione arquivos .xlsx ou .xlsm para começar.",
            style="Subtitle.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(files, textvariable=self.input_summary).grid(
            row=1, column=0, sticky="w", pady=(5, 10)
        )

        list_area = ttk.Frame(files)
        list_area.grid(row=2, column=0, sticky="nsew")
        list_area.columnconfigure(0, weight=1)
        list_area.rowconfigure(0, weight=1)
        self.input_list = self.tk.Listbox(
            list_area,
            exportselection=False,
            activestyle="none",
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            selectmode="browse",
        )
        self.input_list.grid(row=0, column=0, sticky="nsew")
        input_scroll = ttk.Scrollbar(
            list_area,
            orient="vertical",
            command=self.input_list.yview,
        )
        input_scroll.grid(row=0, column=1, sticky="ns")
        self.input_list.configure(yscrollcommand=input_scroll.set)

        self.import_button = self._track_button(
            ttk.Button(
                files,
                text="Adicionar planilhas",
                command=self.import_files,
                style="Primary.TButton",
            )
        )
        self.import_button.grid(row=3, column=0, sticky="ew", pady=(12, 7))

        file_actions = ttk.Frame(files)
        file_actions.grid(row=4, column=0, sticky="ew")
        file_actions.columnconfigure(0, weight=1)
        file_actions.columnconfigure(1, weight=1)
        refresh_button = self._track_button(
            ttk.Button(
                file_actions,
                text="Atualizar lista",
                command=self.refresh_inputs,
                style="Quiet.TButton",
            )
        )
        refresh_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(
            file_actions,
            text="Abrir pasta",
            command=lambda: self.controller.open_directory(
                self.controller.workspace.input_dir
            ),
            style="Quiet.TButton",
        ).grid(row=0, column=1, sticky="ew", padx=(4, 0))

        ttk.Label(
            files,
            text="As ações rápidas usam todas as planilhas desta lista.",
            style="Subtitle.TLabel",
            wraplength=300,
        ).grid(row=5, column=0, sticky="w", pady=(10, 0))

    def _build_request_panel(self, parent: object) -> None:
        from tkinter import ttk

        request = ttk.LabelFrame(
            parent,
            text="2. Descreva o pedido",
            style="Card.TLabelframe",
        )
        request.grid(row=0, column=0, sticky="ew")
        request.columnconfigure(0, weight=1)

        ttk.Label(
            request,
            text=(
                "Use linguagem natural ou escolha uma ação rápida. "
                "Nenhuma alteração ocorre sem prévia."
            ),
            style="Subtitle.TLabel",
            wraplength=680,
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        command_row = ttk.Frame(request)
        command_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        command_row.columnconfigure(0, weight=1)
        self.command_entry = ttk.Entry(
            command_row,
            textvariable=self.command,
            style="Modern.TEntry",
        )
        self.command_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.execute_button = self._track_button(
            ttk.Button(
                command_row,
                text="Executar pedido",
                command=self.execute_command,
                style="Primary.TButton",
            )
        )
        self.execute_button.grid(row=0, column=1, sticky="ew")
        self.voice_button = self._track_button(
            ttk.Button(
                command_row,
                text="Usar voz",
                command=self.capture_voice,
                style="Action.TButton",
            )
        )
        self.voice_button.grid(row=0, column=2, sticky="ew", padx=(8, 0))
        if not self.controller.capabilities.voice_available:
            self.voice_button.state(["disabled"])

        ttk.Label(request, text="Ações rápidas", style="Section.TLabel").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(13, 6)
        )
        quick_actions = ttk.Frame(request)
        quick_actions.grid(row=3, column=0, columnspan=2, sticky="ew")
        for column in range(4):
            quick_actions.columnconfigure(column, weight=1)
        actions = (
            ("Reconhecer", "reconhecer todas"),
            ("Diagnosticar", "diagnosticar todas"),
            ("Limpar e organizar", "limpar e organizar todas"),
            ("Criar relatório", "resumir e criar relatorio de todas"),
        )
        for column, (label, command) in enumerate(actions):
            button = self._track_button(
                ttk.Button(
                    quick_actions,
                    text=label,
                    command=lambda value=command: self._set_and_execute(value),
                    style="Quiet.TButton",
                )
            )
            padding = (0 if column == 0 else 4, 0 if column == 3 else 4)
            button.grid(row=0, column=column, sticky="ew", padx=padding)

    def _build_preview_panel(self, parent: object) -> None:
        from tkinter import ttk

        preview = ttk.LabelFrame(
            parent,
            text="3. Revise a prévia",
            style="Card.TLabelframe",
        )
        preview.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(1, weight=1)

        preview_header = ttk.Frame(preview)
        preview_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        preview_header.columnconfigure(0, weight=1)
        ttk.Label(
            preview_header,
            text="Confira o arquivo, as ações e os avisos antes de confirmar.",
            style="Subtitle.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            preview_header,
            textvariable=self.plan_summary,
            style="Section.TLabel",
        ).grid(row=0, column=1, sticky="e")

        preview_area = ttk.Frame(preview)
        preview_area.grid(row=1, column=0, sticky="nsew")
        preview_area.columnconfigure(0, weight=1)
        preview_area.rowconfigure(0, weight=1)
        self.preview = self.tk.Text(
            preview_area,
            wrap="word",
            state="disabled",
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            padx=12,
            pady=10,
        )
        self.preview.grid(row=0, column=0, sticky="nsew")
        preview_scroll = ttk.Scrollbar(
            preview_area,
            orient="vertical",
            command=self.preview.yview,
        )
        preview_scroll.grid(row=0, column=1, sticky="ns")
        self.preview.configure(yscrollcommand=preview_scroll.set)

        preview_actions = ttk.Frame(preview)
        preview_actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        preview_actions.columnconfigure(0, weight=1)
        preview_actions.columnconfigure(1, weight=1)
        preview_actions.columnconfigure(2, weight=1)
        self.confirm_button = ttk.Button(
            preview_actions,
            text="Confirmar plano",
            command=self.confirm_plan,
            style="Plan.TButton",
        )
        self.confirm_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.cancel_button = ttk.Button(
            preview_actions,
            text="Cancelar plano",
            command=self.cancel_plan,
            style="Action.TButton",
        )
        self.cancel_button.grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(
            preview_actions,
            text="Abrir saída",
            command=lambda: self.controller.open_directory(
                self.controller.workspace.output_dir
            ),
            style="Action.TButton",
        ).grid(row=0, column=2, sticky="ew", padx=(4, 0))

        self._set_preview(
            "A prévia detalhada ou o resultado do diagnóstico aparecerá aqui."
        )
        self._toggle_plan_buttons(False)

    def _build_settings_tab(self, parent: object) -> None:
        from tkinter import ttk

        parent.columnconfigure(0, weight=1)
        ttk.Label(parent, text="Configurações", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            parent,
            text="Preferências salvas para as próximas execuções.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 16))

        settings = ttk.LabelFrame(
            parent,
            text="Preferências do assistente",
            style="Card.TLabelframe",
        )
        settings.grid(row=2, column=0, sticky="new")
        settings.columnconfigure(1, weight=1)

        ttk.Label(settings, text="Nome usado nos arquivos de saída").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Entry(
            settings,
            textvariable=self.candidate_name,
            style="Modern.TEntry",
        ).grid(row=0, column=1, sticky="ew", padx=(14, 0))

        self.native_pivot_check = ttk.Checkbutton(
            settings,
            text="Usar Excel Desktop para criar Tabela Dinâmica nativa",
            variable=self.native_pivot,
        )
        self.native_pivot_check.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(16, 0),
        )
        native_hint = (
            "Disponível neste Windows."
            if self.controller.capabilities.native_excel_available
            else "Indisponível nesta plataforma; será usado o modo compatível."
        )
        ttk.Label(
            settings,
            text=native_hint,
            style="Subtitle.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(3, 0))
        if not self.controller.capabilities.native_excel_available:
            self.native_pivot.set(False)
            self.native_pivot_check.state(["disabled"])

        monitor_row = ttk.Frame(settings)
        monitor_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        monitor_row.columnconfigure(1, weight=1)
        ttk.Label(monitor_row, text="Atualizar lista a cada").grid(
            row=0, column=0, sticky="w"
        )
        self.poll_spinbox = ttk.Spinbox(
            monitor_row,
            from_=0.5,
            to=60,
            increment=0.5,
            textvariable=self.poll_interval,
            width=8,
        )
        self.poll_spinbox.grid(row=0, column=1, sticky="w", padx=(8, 4))
        ttk.Label(monitor_row, text="segundos").grid(row=0, column=2, sticky="w")

        settings_actions = ttk.Frame(settings)
        settings_actions.grid(
            row=4,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(20, 0),
        )
        settings_actions.columnconfigure(0, weight=1)
        settings_actions.columnconfigure(1, weight=1)
        save_button = self._track_button(
            ttk.Button(
                settings_actions,
                text="Salvar configurações",
                command=self.save_settings,
                style="Primary.TButton",
            )
        )
        save_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.monitor_button = self._track_button(
            ttk.Button(
                settings_actions,
                text="Iniciar monitor",
                command=self.toggle_monitor,
                style="Action.TButton",
            )
        )
        self.monitor_button.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        ttk.Label(
            parent,
            text=(
                "O monitor apenas atualiza a lista de planilhas. "
                "Ele nunca confirma nem executa alterações silenciosamente."
            ),
            style="Subtitle.TLabel",
            wraplength=700,
        ).grid(row=3, column=0, sticky="w", pady=(14, 0))

    def _build_status_bar(self, parent: object) -> None:
        from tkinter import ttk

        status_bar = ttk.Frame(parent)
        status_bar.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        status_bar.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(
            status_bar,
            textvariable=self.status,
            style="Status.TLabel",
        )
        self.status_label.grid(row=0, column=0, sticky="w")
        ttk.Label(
            status_bar,
            text="Ctrl/Cmd+O: adicionar · F5: atualizar · Ctrl/Cmd+Enter: executar",
            style="Subtitle.TLabel",
        ).grid(row=0, column=1, sticky="e", padx=(12, 12))
        self.progress = ttk.Progressbar(
            status_bar,
            mode="indeterminate",
            length=120,
        )
        self.progress.grid(row=0, column=2, sticky="e")

    def _bind_shortcuts(self) -> None:
        self.command_entry.bind("<Return>", self._execute_from_event)
        self.root.bind("<Control-Return>", self._execute_from_event)
        self.root.bind("<Command-Return>", self._execute_from_event)
        self.root.bind("<Control-o>", self._import_from_event)
        self.root.bind("<Command-o>", self._import_from_event)
        self.root.bind("<F5>", self._refresh_from_event)
        self.root.bind("<Control-l>", self._focus_command)
        self.root.bind("<Command-l>", self._focus_command)

    def _track_button(self, button: _T) -> _T:
        self.general_buttons.append(button)
        return button

    def _execute_from_event(self, _event: object) -> str:
        self.execute_command()
        return "break"

    def _import_from_event(self, _event: object) -> str:
        self.import_files()
        return "break"

    def _refresh_from_event(self, _event: object) -> str:
        self.refresh_inputs()
        return "break"

    def _focus_command(self, _event: object) -> str:
        self.notebook.select(0)
        self.command_entry.focus_set()
        return "break"

    def _load_settings(self) -> None:
        config = self.controller.load_config()
        self.candidate_name.set(config.candidate_name)
        self.native_pivot.set(
            config.use_native_pivot
            and self.controller.capabilities.native_excel_available
        )
        self.poll_interval.set(str(config.poll_interval_seconds))

    def refresh_inputs(self, *, announce: bool = True) -> None:
        self.input_list.delete(0, "end")
        inputs = self.controller.list_inputs()
        for path in inputs:
            self.input_list.insert("end", path.name)
        count = len(inputs)
        self.input_summary.set(
            "Nenhuma planilha adicionada"
            if count == 0
            else f"{count} planilha(s) pronta(s) para uso"
        )
        if announce:
            self._set_status(
                "Adicione uma planilha para começar"
                if count == 0
                else f"Lista atualizada: {count} planilha(s)",
            )

    def import_files(self) -> None:
        from tkinter import filedialog

        selected = filedialog.askopenfilenames(
            title="Adicionar planilhas",
            filetypes=(("Planilhas Excel", "*.xlsx *.xlsm"),),
        )
        if selected:
            self._submit(
                lambda: self.controller.import_files(
                    tuple(Path(item) for item in selected)
                ),
                self._after_import,
                "Copiando planilhas...",
            )

    def _after_import(self, value: object) -> None:
        imported = tuple(value)  # type: ignore[arg-type]
        self.refresh_inputs(announce=False)
        self._set_status(f"{len(imported)} planilha(s) adicionada(s)", "success")

    def execute_command(self) -> None:
        command = self.command.get().strip()
        if not command:
            self._show_error("Digite ou fale um pedido.")
            return
        self._submit(
            lambda: self.controller.execute(command),
            self._show_result,
            "Interpretando o pedido e preparando a prévia...",
        )

    def _set_and_execute(self, command: str) -> None:
        self.command.set(command)
        self.execute_command()

    def capture_voice(self) -> None:
        self._submit(
            self.controller.recognize_voice,
            lambda result: self._set_voice_text(result.text),
            "Aguardando comando por voz...",
        )

    def _set_voice_text(self, text: str) -> None:
        self.command.set(text)
        self.command_entry.focus_set()
        self._set_status("Comando por voz transcrito. Revise e clique em Executar.")

    def confirm_plan(self) -> None:
        from tkinter import messagebox

        if not self.current_plan_ids:
            return
        plan_count = len(self.current_plan_ids)
        if not messagebox.askyesno(
            "Confirmar alterações",
            (
                f"Executar {plan_count} plano(s) exibido(s)?\n\n"
                "Os resultados serão gravados em novas cópias e os arquivos "
                "originais serão preservados."
            ),
            parent=self.root,
        ):
            return
        plan_ids = self.current_plan_ids
        self._submit(
            lambda: tuple(
                self.controller.confirm_plan(plan_id) for plan_id in plan_ids
            ),
            self._show_plan_results,
            "Aplicando os planos confirmados...",
        )

    def cancel_plan(self) -> None:
        if not self.current_plan_ids:
            return
        plan_ids = self.current_plan_ids
        self._submit(
            lambda: tuple(self.controller.cancel_plan(plan_id) for plan_id in plan_ids),
            self._show_plan_results,
            "Cancelando os planos...",
        )

    def save_settings(self) -> None:
        try:
            interval = float(self.poll_interval.get().replace(",", "."))
        except ValueError:
            self._show_error("O intervalo do monitor deve ser numérico.")
            return
        self._submit(
            lambda: self.controller.save_config(
                candidate_name=self.candidate_name.get(),
                use_native_pivot=self.native_pivot.get(),
                poll_interval_seconds=interval,
            ),
            lambda _: self._set_status("Configurações salvas", "success"),
            "Salvando configurações...",
        )

    def toggle_monitor(self) -> None:
        if self.monitoring:
            self.monitoring = False
            if self.monitor_after_id is not None:
                self.root.after_cancel(self.monitor_after_id)
                self.monitor_after_id = None
            self.monitor_button.configure(text="Iniciar monitor")
            self._set_status("Monitor parado")
            return
        self.monitoring = True
        self.monitor_button.configure(text="Parar monitor")
        self._set_status("Monitor ativo: aguardando novas planilhas", "success")
        self._monitor_tick()

    def _monitor_tick(self) -> None:
        self.monitor_after_id = None
        if not self.monitoring:
            return
        self.refresh_inputs(announce=False)
        self._set_status("Monitor ativo: aguardando novas planilhas", "success")
        milliseconds = int(self.controller.load_config().poll_interval_seconds * 1000)
        self.monitor_after_id = self.root.after(milliseconds, self._monitor_tick)

    def _show_result(self, value: object) -> None:
        result = value
        if not isinstance(result, AssistantResult):
            raise TypeError("Resultado inesperado do assistente.")
        plan_ids, content = self.controller.read_previews(result)
        self.current_plan_ids = plan_ids
        self._toggle_plan_buttons(bool(plan_ids))
        self._set_preview(content)
        self.refresh_inputs(announce=False)
        self._set_status(
            "Prévia pronta: revise e confirme para gerar as cópias"
            if plan_ids
            else "Operação concluída",
            "warning" if plan_ids else "success",
        )

    def _show_plan_results(self, value: object) -> None:
        results = tuple(value)  # type: ignore[arg-type]
        self.current_plan_ids = ()
        self._toggle_plan_buttons(False)
        self._set_preview(
            "\n".join(format_assistant_result(result) for result in results)
        )
        self._set_status(
            "Planos concluídos. Os originais foram preservados.", "success"
        )

    def _submit(
        self,
        operation: Callable[[], _T],
        on_success: Callable[[_T], None] | None,
        status: str,
    ) -> None:
        if self.busy:
            self._set_status("Aguarde a operação atual terminar", "warning")
            return
        self._set_busy(True)
        self._set_status(status)
        future: Future[object] = self.executor.submit(operation)
        self.completed.put((future, on_success))  # type: ignore[arg-type]

    def _poll_completed(self) -> None:
        pending: list[tuple[Future[object], Callable[[object], None] | None]] = []
        while not self.completed.empty():
            future, callback = self.completed.get()
            if not future.done():
                pending.append((future, callback))
                continue
            try:
                result = future.result()
                if callback is not None:
                    callback(result)
            except Exception as error:  # noqa: BLE001 - fronteira visual
                self._show_error(str(error))
                self._stop_monitor_after_error()
            finally:
                self._set_busy(False)
        for item in pending:
            self.completed.put(item)
        self.root.after(100, self._poll_completed)

    def _stop_monitor_after_error(self) -> None:
        if not self.monitoring:
            return
        self.monitoring = False
        if self.monitor_after_id is not None:
            self.root.after_cancel(self.monitor_after_id)
            self.monitor_after_id = None
        self.monitor_button.configure(text="Iniciar monitor")

    def _set_busy(self, active: bool) -> None:
        self.busy = active
        for button in self.general_buttons:
            button.state(["disabled"] if active else ["!disabled"])
        if not active and not self.controller.capabilities.voice_available:
            self.voice_button.state(["disabled"])
        self._toggle_plan_buttons(bool(self.current_plan_ids) and not active)
        if active:
            self.progress.start(12)
        else:
            self.progress.stop()

    def _set_preview(self, content: str) -> None:
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", content)
        self.preview.configure(state="disabled")

    def _toggle_plan_buttons(self, enabled: bool) -> None:
        state = ["!disabled"] if enabled else ["disabled"]
        self.confirm_button.state(state)
        self.cancel_button.state(state)
        count = len(self.current_plan_ids)
        self.plan_summary.set(
            f"{count} plano(s) aguardando" if count else "Nenhum plano aguardando"
        )

    def _set_status(self, message: str, tone: str = "normal") -> None:
        styles = {
            "normal": "Status.TLabel",
            "success": "StatusSuccess.TLabel",
            "warning": "StatusWarning.TLabel",
            "error": "StatusError.TLabel",
        }
        self.status.set(message)
        self.status_label.configure(style=styles.get(tone, "Status.TLabel"))

    def _show_error(self, message: str) -> None:
        from tkinter import messagebox

        self._set_status("Atenção necessária", "error")
        messagebox.showerror("Excel Compras Automation", message, parent=self.root)

    def _close(self) -> None:
        if self.busy:
            from tkinter import messagebox

            messagebox.showinfo(
                "Operação em andamento",
                "Aguarde a operação atual terminar antes de fechar.",
                parent=self.root,
            )
            return
        self.monitoring = False
        if self.monitor_after_id is not None:
            self.root.after_cancel(self.monitor_after_id)
            self.monitor_after_id = None
        self.executor.shutdown(wait=False, cancel_futures=True)
        self.root.destroy()
