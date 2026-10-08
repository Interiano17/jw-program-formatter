"""
Punto de entrada del Meeting Generator.

Proporciona una interfaz gráfica con tkinter para seleccionar
el documento fuente, la plantilla y la ruta donde guardar el resultado.
"""

import logging
import tkinter as tk
from collections.abc import Callable
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from typing import Any

from .error_panel import ErrorPanel
from .generator import generate_document
from .results import GenerationResult
from .settings import AppSettings
from .theme import (
    COLOR_ACCENT,
    COLOR_BACKGROUND,
    COLOR_BORDER,
    COLOR_CARD,
    COLOR_DISABLED,
    COLOR_DISABLED_TEXT,
    COLOR_ERROR,
    COLOR_INFO,
    COLOR_MUTED,
    COLOR_PRIMARY,
    COLOR_PRIMARY_SOFT,
    COLOR_SUCCESS,
    COLOR_WARNING,
)
from .utils import get_default_output_path, setup_logging

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # El arrastrar y soltar es opcional.
    DND_FILES = None
    TkinterDnD = None

logger = logging.getLogger(__name__)

LOGO_PATH = Path(__file__).resolve().parent.parent / "assets" / "logo.png"
HEADER_LOGO_SIZE = 56
NO_FILE_TEXT = "No se ha seleccionado archivo"
WINDOW_WIDTH = 600
COMPACT_HEIGHT = 470
EXPANDED_HEIGHT = 720
MIN_WINDOW_SIZE = (540, 400)
MAX_DIALOG_ERRORS = 8
MAX_DISPLAY_NAME = 34


def _display_name(name: str, limit: int = MAX_DISPLAY_NAME) -> str:
    """Acorta un nombre de archivo largo conservando el final y la extensión."""
    if len(name) <= limit:
        return name
    return f"{name[: limit - 9]}…{name[-8:]}"


class MeetingGeneratorApp:
    """
    Aplicación principal con interfaz gráfica para el Meeting Generator.
    """

    def __init__(self) -> None:
        """Inicializa la aplicación y configura la ventana de tkinter."""
        self.root = self._create_root()
        self.root.title("JW Program Formatter")
        self._set_window_icon()
        self.root.geometry(f"{WINDOW_WIDTH}x{COMPACT_HEIGHT}")
        self.root.minsize(*MIN_WINDOW_SIZE)

        self.source_path: Path | None = None
        self.template_path: Path | None = None
        self._generation_in_progress = False
        self._executor = ThreadPoolExecutor(max_workers=1)
        self._generation_future: Future[GenerationResult] | None = None
        self._generation_poll_id: str | None = None
        self._close_requested = False

        self.settings = AppSettings()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.root.bind("<Control-g>", self._generate_shortcut)
        self.root.bind("<Control-G>", self._generate_shortcut)
        self._restore_saved_template()

    def _create_root(self) -> tk.Tk:
        """Crea la ventana raíz, con soporte de arrastrar y soltar si existe."""
        # className define el WM_CLASS/app_id que GNOME enlaza con el .desktop.
        if TkinterDnD is not None:
            try:
                self._dnd_enabled = True
                return TkinterDnD.Tk(className="generador-s140")
            except (RuntimeError, tk.TclError):
                logger.warning(
                    "tkdnd no está disponible; se desactiva arrastrar y soltar."
                )
        self._dnd_enabled = False
        return tk.Tk(className="generador-s140")

    def _set_window_icon(self) -> None:
        """Asigna el logo como icono de la ventana (Alt+Tab, barra de tareas)."""
        try:
            self._window_icon = tk.PhotoImage(file=str(LOGO_PATH))
            self.root.iconphoto(True, self._window_icon)
        except tk.TclError:
            logger.warning("No se pudo cargar el icono de la ventana: %s", LOGO_PATH)

    def _configure_styles(self) -> None:
        """Define el tema ttk y los estilos propios de la aplicación."""
        base = tkfont.nametofont("TkDefaultFont")
        family = base.actual("family")
        self._fonts = {
            "title": tkfont.Font(family=family, size=18, weight="bold"),
            "subtitle": tkfont.Font(family=family, size=10),
            "caption": tkfont.Font(family=family, size=8, weight="bold"),
            "file": tkfont.Font(family=family, size=10),
            "button": tkfont.Font(family=family, size=10),
            "primary": tkfont.Font(family=family, size=11, weight="bold"),
            "status": tkfont.Font(family=family, size=10),
        }

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "Secondary.TButton",
            font=self._fonts["button"],
            padding=(16, 8),
            background=COLOR_CARD,
            foreground=COLOR_PRIMARY,
            bordercolor=COLOR_PRIMARY,
            lightcolor=COLOR_CARD,
            darkcolor=COLOR_CARD,
            focuscolor=COLOR_CARD,
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Secondary.TButton",
            background=[("active", COLOR_PRIMARY_SOFT), ("disabled", COLOR_CARD)],
            foreground=[("disabled", COLOR_MUTED)],
            bordercolor=[("disabled", COLOR_BORDER)],
            lightcolor=[("active", COLOR_PRIMARY_SOFT)],
            darkcolor=[("active", COLOR_PRIMARY_SOFT)],
        )
        style.configure("Compact.Secondary.TButton", padding=(12, 4))
        style.configure(
            "App.Horizontal.TProgressbar",
            troughcolor=COLOR_BORDER,
            background=COLOR_PRIMARY,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_PRIMARY,
            darkcolor=COLOR_PRIMARY,
            thickness=6,
        )
        style.configure(
            "Primary.TButton",
            font=self._fonts["primary"],
            padding=(16, 12),
            background=COLOR_PRIMARY,
            foreground="white",
            bordercolor=COLOR_PRIMARY,
            lightcolor=COLOR_PRIMARY,
            darkcolor=COLOR_PRIMARY,
            focuscolor=COLOR_PRIMARY,
            relief="flat",
        )
        style.map(
            "Primary.TButton",
            background=[
                ("disabled", COLOR_DISABLED),
                ("pressed", COLOR_ACCENT),
                ("active", COLOR_ACCENT),
            ],
            foreground=[("disabled", COLOR_DISABLED_TEXT)],
            bordercolor=[
                ("disabled", COLOR_DISABLED),
                ("pressed", COLOR_ACCENT),
                ("active", COLOR_ACCENT),
            ],
            lightcolor=[("disabled", COLOR_DISABLED), ("active", COLOR_ACCENT)],
            darkcolor=[("disabled", COLOR_DISABLED), ("active", COLOR_ACCENT)],
        )

    def _build_file_card(
        self,
        parent: tk.Misc,
        caption: str,
        command: Callable[[], None],
        on_drop: Callable[[Path], None],
    ) -> tuple[ttk.Button, tk.Label]:
        """Crea una tarjeta de selección de archivo; devuelve botón y etiqueta."""
        card = tk.Frame(
            parent,
            bg=COLOR_CARD,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_BORDER,
            highlightthickness=1,
        )
        card.pack(fill=tk.X, pady=(0, 12))
        card.columnconfigure(0, weight=1)

        tk.Label(
            card,
            text=caption,
            font=self._fonts["caption"],
            bg=COLOR_CARD,
            fg=COLOR_PRIMARY,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(14, 0))

        file_label = tk.Label(
            card,
            text=NO_FILE_TEXT,
            font=self._fonts["file"],
            bg=COLOR_CARD,
            fg=COLOR_MUTED,
            anchor="w",
            justify="left",
        )
        file_label.grid(row=1, column=0, sticky="w", padx=16, pady=(2, 14))

        button = ttk.Button(
            card,
            text="Examinar…",
            command=command,
            style="Secondary.TButton",
            cursor="hand2",
        )
        button.grid(row=0, column=1, rowspan=2, padx=16, pady=14)
        self._register_drop_target(card, on_drop)
        return button, file_label

    def _register_drop_target(
        self, card: tk.Frame, on_drop: Callable[[Path], None]
    ) -> None:
        """Permite soltar un .docx sobre la tarjeta cuando tkdnd está disponible."""
        if not self._dnd_enabled:
            return

        def highlight(color: str) -> None:
            card.config(highlightbackground=color, highlightcolor=color)

        def handle_enter(event: Any) -> str:
            highlight(COLOR_PRIMARY)
            return str(event.action)

        def handle_leave(event: Any) -> str:
            highlight(COLOR_BORDER)
            return str(event.action)

        def handle_drop(event: Any) -> str:
            highlight(COLOR_BORDER)
            self._handle_drop(str(event.data), on_drop)
            return str(event.action)

        card.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
        card.dnd_bind("<<DropEnter>>", handle_enter)  # type: ignore[attr-defined]
        card.dnd_bind("<<DropLeave>>", handle_leave)  # type: ignore[attr-defined]
        card.dnd_bind("<<Drop>>", handle_drop)  # type: ignore[attr-defined]

    def _handle_drop(self, data: str, assign: Callable[[Path], None]) -> None:
        """Asigna el primer .docx soltado sobre una tarjeta."""
        if self._generation_in_progress:
            return
        dropped = [Path(item) for item in self.root.tk.splitlist(data)]
        documents = [path for path in dropped if path.suffix.lower() == ".docx"]
        if not documents:
            self.status_label.config(
                text="Solo se aceptan archivos .docx.", fg=COLOR_WARNING
            )
            return
        assign(documents[0])

    def _build_ui(self) -> None:
        """Construye la interfaz de usuario."""
        self._configure_styles()
        self.root.configure(bg=COLOR_BACKGROUND)

        content = tk.Frame(self.root, bg=COLOR_BACKGROUND)
        content.pack(fill=tk.BOTH, expand=True, padx=32, pady=28)

        header = tk.Frame(content, bg=COLOR_BACKGROUND)
        header.pack(fill=tk.X, pady=(0, 24))

        header_logo = self._load_header_logo()
        if header_logo is not None:
            tk.Label(header, image=header_logo, bg=COLOR_BACKGROUND).pack(
                side=tk.LEFT, padx=(0, 16)
            )

        titles = tk.Frame(header, bg=COLOR_BACKGROUND)
        titles.pack(side=tk.LEFT, fill=tk.X)
        tk.Label(
            titles,
            text="JW Program Formatter",
            font=self._fonts["title"],
            bg=COLOR_BACKGROUND,
            fg=COLOR_ACCENT,
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            titles,
            text="Convierte el programa de la reunión en hojas S-140.",
            font=self._fonts["subtitle"],
            bg=COLOR_BACKGROUND,
            fg=COLOR_MUTED,
            anchor="w",
        ).pack(anchor="w")

        self.source_button, self.source_label = self._build_file_card(
            content,
            "1 · DOCUMENTO DEL PROGRAMA",
            self._select_source_file,
            self._set_source,
        )
        self.template_button, self.template_label = self._build_file_card(
            content,
            "2 · PLANTILLA S-140",
            self._select_template_file,
            self._set_template,
        )

        self.generate_button = ttk.Button(
            content,
            text="Generar y guardar como…",
            command=self._generate_document,
            style="Primary.TButton",
            cursor="hand2",
            state=tk.DISABLED,
        )
        self.generate_button.pack(fill=tk.X, pady=(12, 0))

        # El contenedor reserva el espacio para que el layout no salte al generar.
        self._progress_holder = tk.Frame(content, bg=COLOR_BACKGROUND, height=16)
        self._progress_holder.pack(fill=tk.X)
        self._progress_holder.pack_propagate(False)
        self.progress = ttk.Progressbar(
            self._progress_holder,
            mode="indeterminate",
            style="App.Horizontal.TProgressbar",
        )

        self.status_label = tk.Label(
            content,
            text="Seleccione ambos archivos para continuar.",
            font=self._fonts["status"],
            bg=COLOR_BACKGROUND,
            fg=COLOR_INFO,
            anchor="w",
            justify="left",
            wraplength=500,
        )
        self.status_label.pack(fill=tk.X, pady=(14, 0))

        self.error_panel = ErrorPanel(
            content,
            background=COLOR_BACKGROUND,
            family=self._fonts["status"].actual("family"),
        )
        content.bind("<Configure>", self._adjust_wraplengths)

    def _adjust_wraplengths(self, event: tk.Event) -> None:
        """Ajusta el salto de línea del estado al ancho disponible."""
        self.status_label.config(wraplength=max(200, event.width))

    def _load_header_logo(self) -> tk.PhotoImage | None:
        """Carga el logo reducido para el encabezado, o None si no está disponible."""
        try:
            logo = tk.PhotoImage(file=str(LOGO_PATH))
        except tk.TclError:
            return None
        factor = max(1, logo.height() // HEADER_LOGO_SIZE)
        self._header_logo = logo.subsample(factor, factor)
        return self._header_logo

    def _select_source_file(self) -> None:
        """Abre el diálogo para seleccionar el documento fuente del programa."""
        if self._generation_in_progress:
            return

        file_path = filedialog.askopenfilename(
            title="Seleccionar Documento del Programa",
            initialdir=self._initial_directory("source_dir"),
            filetypes=[
                ("Documentos Word", "*.docx"),
                ("Todos los archivos", "*.*"),
            ],
        )

        if file_path:
            self._set_source(Path(file_path))
        else:
            self._update_generate_button()

    def _select_template_file(self) -> None:
        """Abre el diálogo para seleccionar la plantilla S-140."""
        if self._generation_in_progress:
            return

        file_path = filedialog.askopenfilename(
            title="Seleccionar Plantilla S-140",
            initialdir=self._initial_directory("template_dir"),
            filetypes=[
                ("Documentos Word", "*.docx"),
                ("Todos los archivos", "*.*"),
            ],
        )

        if file_path:
            self._set_template(Path(file_path))
        else:
            self._update_generate_button()

    def _initial_directory(self, key: str) -> str | None:
        """Devuelve la última carpeta usada para ese tipo de archivo."""
        directory = self.settings.get_path(key)
        return str(directory) if directory is not None and directory.is_dir() else None

    def _set_source(self, path: Path) -> None:
        """Registra el documento fuente elegido y recuerda su carpeta."""
        self.source_path = path
        self.source_label.config(text=f"✓ {_display_name(path.name)}", fg=COLOR_SUCCESS)
        self.settings.set_path("source_dir", path.parent)
        logger.info("Documento fuente seleccionado.")
        self._update_generate_button()

    def _set_template(self, path: Path) -> None:
        """Registra la plantilla elegida y la recuerda para la próxima sesión."""
        self.template_path = path
        self.template_label.config(
            text=f"✓ {_display_name(path.name)}", fg=COLOR_SUCCESS
        )
        self.settings.set_path("template_dir", path.parent)
        self.settings.set_path("template_file", path)
        logger.info("Plantilla seleccionada.")
        self._update_generate_button()

    def _restore_saved_template(self) -> None:
        """Vuelve a cargar la plantilla de la sesión anterior si aún existe."""
        saved = self.settings.get_path("template_file")
        if saved is not None and saved.is_file():
            self._set_template(saved)

    def _generate_shortcut(self, _event: tk.Event) -> str:
        """Atajo Ctrl+G: genera solo si el botón está habilitado."""
        if (
            not self._generation_in_progress
            and self.source_path is not None
            and self.template_path is not None
        ):
            self._generate_document()
        return "break"

    def _update_generate_button(self) -> None:
        """Actualiza el estado del botón Generar según los archivos seleccionados."""
        if self._generation_in_progress:
            self.generate_button.config(state=tk.DISABLED)
        elif self.source_path and self.template_path:
            self.generate_button.config(state=tk.NORMAL)
            self.status_label.config(
                text="Listo para generar el documento.",
                fg=COLOR_SUCCESS,
            )
        else:
            self.generate_button.config(state=tk.DISABLED)
            self.status_label.config(
                text="Seleccione ambos archivos para continuar.",
                fg=COLOR_INFO,
            )

    def _set_generation_in_progress(self, in_progress: bool) -> None:
        """Bloquea o restaura todos los controles que pueden cambiar las rutas."""
        self._generation_in_progress = in_progress
        if in_progress:
            self.progress.pack(fill=tk.X, pady=(10, 0))
            self.progress.start(12)
            self.source_button.config(state=tk.DISABLED)
            self.template_button.config(state=tk.DISABLED)
            self.generate_button.config(state=tk.DISABLED)
            return

        self.progress.stop()
        self.progress.pack_forget()
        self.source_button.config(state=tk.NORMAL)
        self.template_button.config(state=tk.NORMAL)
        if self.source_path is not None and self.template_path is not None:
            self.generate_button.config(state=tk.NORMAL)
        else:
            self.generate_button.config(state=tk.DISABLED)

    @staticmethod
    def _format_error_details(errors: tuple[str, ...]) -> str:
        """Resume los errores para el diálogo; el detalle completo va al panel."""
        lines = [f"• {error}" for error in errors[:MAX_DIALOG_ERRORS]]
        hidden = len(errors) - MAX_DIALOG_ERRORS
        if hidden > 0:
            lines.append(f"• … y {hidden} más (vea el panel de detalles)")
        return "\n".join(lines)

    def _show_error_details(self, errors: tuple[str, ...] | list[str]) -> None:
        """Muestra todos los errores agrupados en la ventana principal."""
        self.error_panel.show(errors)
        self.root.geometry(f"{WINDOW_WIDTH}x{EXPANDED_HEIGHT}")

    def _hide_error_details(self) -> None:
        """Oculta el panel de detalles y devuelve la ventana a su tamaño base."""
        if self.error_panel.visible:
            self.error_panel.hide()
            self.root.geometry(f"{WINDOW_WIDTH}x{COMPACT_HEIGHT}")

    def _select_output_file(
        self,
        source_path: Path,
        template_path: Path,
    ) -> Path | None:
        """Solicita una ruta de salida y confirma el reemplazo si ya existe."""
        suggested_path = get_default_output_path(source_path, template_path)
        selected_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Guardar documento S-140 como",
            initialdir=str(suggested_path.parent),
            initialfile=suggested_path.name,
            defaultextension=".docx",
            filetypes=[
                ("Documentos Word", "*.docx"),
                ("Todos los archivos", "*.*"),
            ],
            confirmoverwrite=False,
        )
        if not selected_path:
            return None

        output_path = Path(selected_path)
        try:
            resolved_output = output_path.resolve()
            protected_paths = {
                source_path.resolve(): "documento fuente",
                template_path.resolve(): "plantilla",
            }
        except (OSError, RuntimeError) as exc:
            messagebox.showerror(
                "Ruta de salida inválida",
                f"No se pudo validar la ruta de salida:\n\n{exc}",
                parent=self.root,
            )
            return None

        collision = protected_paths.get(resolved_output)
        if collision is not None:
            messagebox.showerror(
                "Ruta de salida inválida",
                (
                    f"La salida coincide con {collision}.\n\n"
                    "Elija un archivo diferente para no sobrescribirlo."
                ),
                parent=self.root,
            )
            return None

        if output_path.exists() and not messagebox.askyesno(
            "Confirmar reemplazo",
            (f"El archivo '{output_path.name}' ya existe.\n\n¿Desea reemplazarlo?"),
            parent=self.root,
        ):
            return None
        return output_path

    def _generate_document(self) -> None:
        """Selecciona la salida e inicia la generación en el único trabajador."""
        if self._close_requested:
            return
        if self._generation_in_progress:
            logger.warning("Se ignoró una solicitud de generación reentrante")
            return

        source_path = self.source_path
        template_path = self.template_path
        if source_path is None or template_path is None:
            messagebox.showerror(
                "Error",
                "Debe seleccionar ambos archivos antes de generar.",
            )
            return

        self._set_generation_in_progress(True)
        self._hide_error_details()
        try:
            self.status_label.config(
                text="Seleccione dónde guardar el documento...",
                fg=COLOR_INFO,
            )
            self.root.update_idletasks()

            output_path = self._select_output_file(source_path, template_path)
            if output_path is None or self._close_requested:
                if not self._close_requested:
                    self.status_label.config(
                        text="Generación cancelada. No se modificó ningún archivo.",
                        fg=COLOR_INFO,
                    )
                return

            self.status_label.config(
                text="Validando y generando el documento...",
                fg=COLOR_WARNING,
            )
            output_existed_before = output_path.is_file()

            logger.info("Iniciando generación fail-closed...")
            # El trabajador no accede a Tkinter, ni siquiera para programar callbacks.
            self._generation_poll_id = self.root.after(
                100, self._poll_generation, output_existed_before
            )
            self._generation_future = self._executor.submit(
                generate_document,
                source_path,
                template_path,
                output_path,
            )
        except Exception as exc:
            self._show_generation_error(exc)
        finally:
            if self._generation_future is None:
                self._finish_generation()

    def _poll_generation(self, output_existed_before: bool) -> None:
        """Consulta el resultado sin bloquear el hilo principal de Tkinter."""
        self._generation_poll_id = None
        future = self._generation_future
        if future is None:
            return
        if not future.done():
            self._generation_poll_id = self.root.after(
                100, self._poll_generation, output_existed_before
            )
            return

        try:
            result = future.result()
            if not self._close_requested:
                self._show_generation_result(result, output_existed_before)
        except Exception as exc:
            self._show_generation_error(exc)
        finally:
            self._generation_future = None
            self._finish_generation()

    def _show_generation_result(
        self, result: GenerationResult, output_existed_before: bool
    ) -> None:
        """Presenta el resultado confirmado desde el hilo de la interfaz."""
        generated_path = result.output_path
        succeeded = result.succeeded
        if result.published and not succeeded:
            details = self._format_error_details(result.errors)
            publication_note = (
                "El archivo anterior fue reemplazado por esta ejecución."
                if output_existed_before
                else "El archivo fue creado por esta ejecución."
            )
            error_message = (
                f"El archivo se publicó en:\n\n{generated_path}\n\n"
                "La ejecución terminó con un error posterior a la publicación.\n\n"
                f"Semanas escritas: {result.written_weeks}\n"
                f"Semanas omitidas: {result.omitted_weeks}\n\n"
                "Errores:\n"
                f"{details or '• No se pudo verificar el archivo publicado'}\n\n"
                f"{publication_note}"
            )
            try:
                logger.error(
                    "Fallo posterior a la publicación: escritas=%d omitidas=%d errores=%s",
                    result.written_weeks,
                    result.omitted_weeks,
                    result.errors or ("No se pudo verificar el archivo publicado",),
                )
            except Exception as exc:
                error_message += f"\n\nNo se pudo registrar el error en el log: {exc}"
            self._show_error_details(result.errors)
            messagebox.showerror("Error posterior a la publicación", error_message)
            self.status_label.config(
                text=(
                    "Archivo publicado con errores. "
                    f"Semanas escritas: {result.written_weeks}. "
                    f"Semanas omitidas: {result.omitted_weeks}."
                ),
                fg=COLOR_ERROR,
            )
            return

        if not succeeded or generated_path is None:
            details = self._format_error_details(result.errors)
            previous_file_note = (
                "\n\nYa existía un archivo en la ruta de salida; "
                "no corresponde a esta ejecución."
                if output_existed_before
                else ""
            )
            error_message = (
                "No se confirmó un archivo nuevo para esta ejecución.\n\n"
                f"Semanas escritas: {result.written_weeks}\n"
                f"Semanas omitidas: {result.omitted_weeks}\n\n"
                f"Errores:\n{details or '• Error no especificado'}"
                f"{previous_file_note}"
            )
            logger.error(
                "Generación fallida: escritas=%d omitidas=%d errores=%s",
                result.written_weeks,
                result.omitted_weeks,
                result.errors,
            )
            self._show_error_details(result.errors)
            messagebox.showerror("Error de generación", error_message)
            self.status_label.config(
                text=(
                    "No se generó un archivo nuevo. "
                    f"Semanas omitidas: {result.omitted_weeks}."
                ),
                fg=COLOR_ERROR,
            )
            return

        self.status_label.config(
            text=f"Documento generado exitosamente: {generated_path.name}",
            fg=COLOR_SUCCESS,
        )

        messagebox.showinfo(
            "Éxito",
            f"Documento generado correctamente:\n\n"
            f"📄 {generated_path.name}\n"
            f"📁 {generated_path.parent}\n\n"
            f"Se escribieron {result.written_weeks} semanas.\n"
            f"Se omitieron {result.omitted_weeks} semanas.\n"
            "Consulte el registro de la aplicación para más detalles.",
        )

    def _show_generation_error(self, exc: Exception) -> None:
        """Registra una excepción y la muestra si la ventana sigue abierta."""
        logger.error("Error en la generación: %s", exc, exc_info=True)
        if self._close_requested:
            return
        error_message = (
            f"Error al generar el documento:\n\n{str(exc)}\n\n"
            "Consulte el registro de la aplicación para más detalles."
        )
        self._show_error_details([str(exc)])
        messagebox.showerror("Error", error_message)
        self.status_label.config(
            text=f"Error: {str(exc)[:80]}",
            fg=COLOR_ERROR,
        )

    def _finish_generation(self) -> None:
        """Limpia la consulta pendiente y restaura los controles en todos los casos."""
        if self._generation_poll_id is not None:
            self.root.after_cancel(self._generation_poll_id)
            self._generation_poll_id = None
        self._set_generation_in_progress(False)
        if self._close_requested:
            self._close()

    def _close(self) -> None:
        """Espera la escritura en curso antes de cerrar la ventana y el trabajador."""
        self._close_requested = True
        if self._generation_in_progress:
            self.status_label.config(
                text="Cerrando al terminar la generación...",
                fg=COLOR_WARNING,
            )
            return
        self._executor.shutdown(wait=False, cancel_futures=True)
        self.root.destroy()

    def run(self) -> None:
        """Ejecuta la aplicación."""
        self.root.mainloop()


def main() -> None:
    """Función principal que inicia la aplicación."""
    file_logging_available = setup_logging()

    logger.info("=" * 60)
    logger.info("Meeting Generator - Iniciando aplicación")
    logger.info("=" * 60)

    app = MeetingGeneratorApp()
    if not file_logging_available:
        messagebox.showwarning(
            "Archivo de log no disponible",
            "No se pudo abrir el archivo de log solicitado.\n\n"
            "La aplicación continuará y los registros se escribirán en consola.",
            parent=app.root,
        )
    app.run()


if __name__ == "__main__":
    main()
