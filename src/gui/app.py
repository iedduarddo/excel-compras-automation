"""Janela desktop multiplataforma do Excel Compras Automation."""

from __future__ import annotations

import queue
import sys
from collections.abc import Callable
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import TypeVar

from src.assistant.service import AssistantResult, format_assistant_result
from src.core.exceptions import AutomationError
from src.gui.controller import DesktopController

_T = TypeVar("_T")


def run_desktop_app(root: Path | None = None) -> int:
    """Abre a janela somente quando o modo gráfico é solicitado."""

    import tkinter as tk
    from tkinter import messagebox

    try:
        root_window = tk.Tk()
    except tk.TclError as error:
        print(f"Não foi possível abrir a interface gráfica: {error}", file=sys.stderr)
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
    """Interface orientada a tarefas, com execução em segundo plano."""

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
        self.busy = False

        self.root.title("Excel Compras Automation")
        initial_width = min(1080, self.root.winfo_screenwidth() - 40)
        initial_height = min(720, self.root.winfo_screenheight() - 90)
        self.root.geometry(f"{initial_width}x{initial_height}")
        self.root.minsize(900, 620)
        self.root.protocol("WM_DELETE_WINDOW", self._close)

        self.status = tk.StringVar(value="Pronto")
        self.command = tk.StringVar(value="reconhecer todas")
        self.candidate_name = tk.StringVar()
        self.native_pivot = tk.BooleanVar(value=True)
        self.poll_interval = tk.StringVar(value="2")
        self._build()
        self._load_settings()
        self.refresh_inputs()
        self.root.after(100, self._poll_completed)

    def _build(self) -> None:
        from tkinter import ttk

        container = ttk.Frame(self.root, padding=14)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=2)
        container.columnconfigure(1, weight=3)
        container.rowconfigure(1, weight=1)

        title = ttk.Label(
            container,
            text="Excel Compras Automation",
            font=("Segoe UI", 18, "bold"),
        )
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        files = ttk.LabelFrame(container, text="1. Planilhas de entrada", padding=10)
        files.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        files.rowconfigure(0, weight=1)
        files.columnconfigure(0, weight=1)
        self.input_list = self.tk.Listbox(files, exportselection=False)
        self.input_list.grid(row=0, column=0, columnspan=3, sticky="nsew")
        ttk.Button(files, text="Adicionar", command=self.import_files).grid(
            row=1, column=0, sticky="ew", pady=(8, 0)
        )
        ttk.Button(files, text="Atualizar", command=self.refresh_inputs).grid(
            row=1, column=1, sticky="ew", padx=6, pady=(8, 0)
        )
        ttk.Button(
            files,
            text="Abrir pasta",
            command=lambda: self.controller.open_directory(
                self.controller.workspace.input_dir
            ),
        ).grid(row=1, column=2, sticky="ew", pady=(8, 0))

        right = ttk.Frame(container)
        right.grid(row=1, column=1, sticky="nsew", padx=(8, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        request = ttk.LabelFrame(right, text="2. Pedido", padding=10)
        request.grid(row=0, column=0, sticky="ew")
        request.columnconfigure(0, weight=1)
        ttk.Entry(request, textvariable=self.command).grid(
            row=0, column=0, columnspan=4, sticky="ew"
        )
        ttk.Button(request, text="Executar", command=self.execute_command).grid(
            row=1, column=0, sticky="ew", pady=(8, 0)
        )
        self.voice_button = ttk.Button(request, text="Voz", command=self.capture_voice)
        self.voice_button.grid(row=1, column=1, sticky="ew", padx=6, pady=(8, 0))
        if not self.controller.capabilities.voice_available:
            self.voice_button.state(["disabled"])
        ttk.Button(
            request,
            text="Reconhecer",
            command=lambda: self._set_and_execute("reconhecer todas"),
        ).grid(row=1, column=2, sticky="ew", pady=(8, 0))
        ttk.Button(
            request,
            text="Diagnosticar",
            command=lambda: self._set_and_execute("diagnosticar todas"),
        ).grid(row=1, column=3, sticky="ew", padx=(6, 0), pady=(8, 0))
        ttk.Button(
            request,
            text="Limpar e organizar",
            command=lambda: self._set_and_execute("limpar e organizar todas"),
        ).grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(
            request,
            text="Criar relatório",
            command=lambda: self._set_and_execute("resumir e criar relatorio de todas"),
        ).grid(row=2, column=2, columnspan=2, sticky="ew", padx=(6, 0), pady=(6, 0))

        preview = ttk.LabelFrame(right, text="3. Prévia e resultado", padding=10)
        preview.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        preview.rowconfigure(0, weight=1)
        preview.columnconfigure(0, weight=1)
        self.preview = self.tk.Text(preview, wrap="word", state="disabled")
        self.preview.grid(row=0, column=0, columnspan=3, sticky="nsew")
        self.confirm_button = ttk.Button(
            preview, text="Confirmar plano", command=self.confirm_plan
        )
        self.confirm_button.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self.cancel_button = ttk.Button(
            preview, text="Cancelar plano", command=self.cancel_plan
        )
        self.cancel_button.grid(row=1, column=1, sticky="ew", padx=6, pady=(8, 0))
        ttk.Button(
            preview,
            text="Abrir saída",
            command=lambda: self.controller.open_directory(
                self.controller.workspace.output_dir
            ),
        ).grid(row=1, column=2, sticky="ew", pady=(8, 0))
        self._toggle_plan_buttons(False)

        settings = ttk.LabelFrame(container, text="Configurações", padding=10)
        settings.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        settings.columnconfigure(1, weight=1)
        ttk.Label(settings, text="Nome:").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings, textvariable=self.candidate_name).grid(
            row=0, column=1, sticky="ew", padx=6
        )
        self.native_pivot_check = ttk.Checkbutton(
            settings, text="Usar Excel Desktop no Windows", variable=self.native_pivot
        )
        self.native_pivot_check.grid(row=0, column=2, sticky="w")
        if not self.controller.capabilities.native_excel_available:
            self.native_pivot.set(False)
            self.native_pivot_check.state(["disabled"])
        ttk.Label(settings, text="Monitor (s):").grid(row=0, column=3, padx=(12, 0))
        ttk.Entry(settings, textvariable=self.poll_interval, width=6).grid(
            row=0, column=4
        )
        ttk.Button(settings, text="Salvar", command=self.save_settings).grid(
            row=0, column=5, padx=(8, 0)
        )
        self.monitor_button = ttk.Button(
            settings, text="Iniciar monitor", command=self.toggle_monitor
        )
        self.monitor_button.grid(row=0, column=6, padx=(8, 0))

        ttk.Label(container, textvariable=self.status, anchor="w").grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )

    def _load_settings(self) -> None:
        config = self.controller.load_config()
        self.candidate_name.set(config.candidate_name)
        self.native_pivot.set(
            config.use_native_pivot
            and self.controller.capabilities.native_excel_available
        )
        self.poll_interval.set(str(config.poll_interval_seconds))

    def refresh_inputs(self) -> None:
        self.input_list.delete(0, "end")
        for path in self.controller.list_inputs():
            self.input_list.insert("end", path.name)
        self.status.set(f"{self.input_list.size()} planilha(s) na entrada")

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
                lambda _: self.refresh_inputs(),
                "Copiando planilhas...",
            )

    def execute_command(self) -> None:
        command = self.command.get().strip()
        if not command:
            self._show_error("Digite ou fale um pedido.")
            return
        self._submit(
            lambda: self.controller.execute(command),
            self._show_result,
            "Executando pedido...",
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
        self.status.set("Comando por voz revisado. Clique em Executar.")

    def confirm_plan(self) -> None:
        from tkinter import messagebox

        if not self.current_plan_ids:
            return
        plan_count = len(self.current_plan_ids)
        if not messagebox.askyesno(
            "Confirmar alterações",
            f"Executar {plan_count} plano(s) exibido(s) em novas cópias?",
            parent=self.root,
        ):
            return
        plan_ids = self.current_plan_ids
        self._submit(
            lambda: tuple(
                self.controller.confirm_plan(plan_id) for plan_id in plan_ids
            ),
            self._show_plan_results,
            "Aplicando planos confirmados...",
        )

    def cancel_plan(self) -> None:
        if not self.current_plan_ids:
            return
        plan_ids = self.current_plan_ids
        self._submit(
            lambda: tuple(self.controller.cancel_plan(plan_id) for plan_id in plan_ids),
            self._show_plan_results,
            "Cancelando planos...",
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
            lambda _: self.status.set("Configurações salvas"),
            "Salvando configurações...",
        )

    def toggle_monitor(self) -> None:
        self.monitoring = not self.monitoring
        self.monitor_button.configure(
            text="Parar monitor" if self.monitoring else "Iniciar monitor"
        )
        if self.monitoring:
            self.status.set("Monitor ativo")
            self._monitor_tick()
        else:
            self.status.set("Monitor parado")

    def _monitor_tick(self) -> None:
        if not self.monitoring:
            return
        self.refresh_inputs()
        self.status.set("Monitor ativo: aguardando novas planilhas")
        milliseconds = int(self.controller.load_config().poll_interval_seconds * 1000)
        self.root.after(milliseconds, self._monitor_tick)

    def _show_result(self, value: object) -> None:
        result = value
        if not isinstance(result, AssistantResult):
            raise TypeError("Resultado inesperado do assistente.")
        plan_ids, content = self.controller.read_previews(result)
        self.current_plan_ids = plan_ids
        self._toggle_plan_buttons(bool(plan_ids))
        self._set_preview(content)
        self.status.set("Prévia aguardando confirmação" if plan_ids else "Concluído")
        self.refresh_inputs()

    def _show_plan_results(self, value: object) -> None:
        results = tuple(value)  # type: ignore[arg-type]
        self.current_plan_ids = ()
        self._toggle_plan_buttons(False)
        self._set_preview(
            "\n".join(format_assistant_result(result) for result in results)
        )
        self.status.set("Planos concluídos")

    def _submit(
        self,
        operation: Callable[[], _T],
        on_success: Callable[[_T], None] | None,
        status: str,
    ) -> None:
        if self.busy:
            self.status.set("Aguarde a operação atual terminar")
            return
        self.busy = True
        self.status.set(status)
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
                if self.monitoring:
                    self.monitoring = False
                    self.monitor_button.configure(text="Iniciar monitor")
            finally:
                self.busy = False
        for item in pending:
            self.completed.put(item)
        self.root.after(100, self._poll_completed)

    def _set_preview(self, content: str) -> None:
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", content)
        self.preview.configure(state="disabled")

    def _toggle_plan_buttons(self, enabled: bool) -> None:
        state = ["!disabled"] if enabled else ["disabled"]
        self.confirm_button.state(state)
        self.cancel_button.state(state)

    def _show_error(self, message: str) -> None:
        from tkinter import messagebox

        self.status.set("Atenção necessária")
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
        self.executor.shutdown(wait=False, cancel_futures=True)
        self.root.destroy()
