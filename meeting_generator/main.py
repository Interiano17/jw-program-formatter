"""
Punto de entrada del Meeting Generator.

Proporciona una interfaz gráfica con tkinter para seleccionar
el documento fuente, la plantilla y la ruta donde guardar el resultado.
"""

import logging
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from .generator import generate_document
from .utils import get_default_output_path, setup_logging

logger = logging.getLogger(__name__)


class MeetingGeneratorApp:
    """
    Aplicación principal con interfaz gráfica para el Meeting Generator.
    """

    def __init__(self) -> None:
        """Inicializa la aplicación y configura la ventana de tkinter."""
        self.root = tk.Tk()
        self.root.title("Meeting Generator - S-140")
        self.root.geometry("600x450")
        self.root.resizable(False, False)

        # Variables de estado
        self.source_path: Path | None = None
        self.template_path: Path | None = None
        self._generation_in_progress = False

        self._build_ui()

    def _build_ui(self) -> None:
        """Construye la interfaz de usuario."""
        # Título
        title_label = tk.Label(
            self.root,
            text="Generador de Programa S-140",
            font=("Arial", 16, "bold"),
            pady=20,
        )
        title_label.pack()

        # Descripción
        description = tk.Label(
            self.root,
            text=(
                "Seleccione el documento del programa y la plantilla S-140\n"
                "para generar automáticamente el documento completado."
            ),
            font=("Arial", 10),
            pady=10,
        )
        description.pack()

        # Frame para la selección de archivos
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=20, padx=40, fill=tk.X)

        # Botón para seleccionar documento fuente
        self.source_button = tk.Button(
            file_frame,
            text="1. Seleccionar Documento del Programa",
            command=self._select_source_file,
            width=40,
            height=2,
            bg="#e0e0e0",
        )
        self.source_button.pack(pady=5)

        self.source_label = tk.Label(
            file_frame,
            text="No se ha seleccionado archivo",
            font=("Arial", 9),
            fg="gray",
        )
        self.source_label.pack(pady=2)

        # Botón para seleccionar plantilla
        self.template_button = tk.Button(
            file_frame,
            text="2. Seleccionar Plantilla S-140",
            command=self._select_template_file,
            width=40,
            height=2,
            bg="#e0e0e0",
        )
        self.template_button.pack(pady=5)

        self.template_label = tk.Label(
            file_frame,
            text="No se ha seleccionado archivo",
            font=("Arial", 9),
            fg="gray",
        )
        self.template_label.pack(pady=2)

        # Frame para el botón de generar
        generate_frame = tk.Frame(self.root)
        generate_frame.pack(pady=20)

        self.generate_button = tk.Button(
            generate_frame,
            text="3. Generar y guardar como...",
            command=self._generate_document,
            width=35,
            height=3,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            state=tk.DISABLED,
        )
        self.generate_button.pack()

        # Barra de estado
        self.status_label = tk.Label(
            self.root,
            text="Listo para comenzar.",
            font=("Arial", 9),
            fg="blue",
            pady=10,
        )
        self.status_label.pack()

        # Información de salida
        output_frame = tk.Frame(self.root)
        output_frame.pack(pady=5)

        output_info = tk.Label(
            output_frame,
            text=(
                "Al generar podrá elegir el nombre y la ubicación\n"
                "del archivo de salida."
            ),
            font=("Arial", 8),
            fg="gray",
        )
        output_info.pack()

    def _select_source_file(self) -> None:
        """Abre el diálogo para seleccionar el documento fuente del programa."""
        if self._generation_in_progress:
            return

        file_path = filedialog.askopenfilename(
            title="Seleccionar Documento del Programa",
            filetypes=[
                ("Documentos Word", "*.docx"),
                ("Todos los archivos", "*.*"),
            ],
        )

        if file_path:
            self.source_path = Path(file_path)
            self.source_label.config(
                text=f"✓ {self.source_path.name}",
                fg="green",
            )
            logger.info("Documento fuente seleccionado: %s", self.source_path)
        else:
            self.source_label.config(
                text="No se ha seleccionado archivo",
                fg="gray",
            )

        self._update_generate_button()

    def _select_template_file(self) -> None:
        """Abre el diálogo para seleccionar la plantilla S-140."""
        if self._generation_in_progress:
            return

        file_path = filedialog.askopenfilename(
            title="Seleccionar Plantilla S-140",
            filetypes=[
                ("Documentos Word", "*.docx"),
                ("Todos los archivos", "*.*"),
            ],
        )

        if file_path:
            self.template_path = Path(file_path)
            self.template_label.config(
                text=f"✓ {self.template_path.name}",
                fg="green",
            )
            logger.info("Plantilla seleccionada: %s", self.template_path)
        else:
            self.template_label.config(
                text="No se ha seleccionado archivo",
                fg="gray",
            )

        self._update_generate_button()

    def _update_generate_button(self) -> None:
        """Actualiza el estado del botón Generar según los archivos seleccionados."""
        if self._generation_in_progress:
            self.generate_button.config(state=tk.DISABLED, bg="#4CAF50")
        elif self.source_path and self.template_path:
            self.generate_button.config(state=tk.NORMAL, bg="#4CAF50")
            self.status_label.config(
                text="Listo para generar el documento.",
                fg="green",
            )
        else:
            self.generate_button.config(state=tk.DISABLED, bg="#4CAF50")
            self.status_label.config(
                text="Seleccione ambos archivos para continuar.",
                fg="blue",
            )

    def _set_generation_in_progress(self, in_progress: bool) -> None:
        """Bloquea o restaura todos los controles que pueden cambiar las rutas."""
        self._generation_in_progress = in_progress
        if in_progress:
            self.source_button.config(state=tk.DISABLED)
            self.template_button.config(state=tk.DISABLED)
            self.generate_button.config(state=tk.DISABLED, bg="#4CAF50")
            return

        self.source_button.config(state=tk.NORMAL)
        self.template_button.config(state=tk.NORMAL)
        if self.source_path is not None and self.template_path is not None:
            self.generate_button.config(state=tk.NORMAL, bg="#4CAF50")
        else:
            self.generate_button.config(state=tk.DISABLED, bg="#4CAF50")

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
            (
                f"El archivo '{output_path.name}' ya existe.\n\n"
                "¿Desea reemplazarlo?"
            ),
            parent=self.root,
        ):
            return None
        return output_path

    def _generate_document(self) -> None:
        """
        Ejecuta el proceso de generación del documento completo.

        Realiza el parsing del documento fuente, rellena la plantilla
        y guarda el archivo de salida.
        """
        if getattr(self, "_generation_in_progress", False):
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
        try:
            self.status_label.config(
                text="Seleccione dónde guardar el documento...",
                fg="blue",
            )
            self.root.update_idletasks()

            output_path = self._select_output_file(source_path, template_path)
            if output_path is None:
                self.status_label.config(
                    text="Generación cancelada. No se modificó ningún archivo.",
                    fg="blue",
                )
                return

            self.status_label.config(
                text="Validando y generando el documento...",
                fg="orange",
            )
            output_existed_before = output_path.is_file()

            logger.info("Iniciando generación fail-closed...")
            result = generate_document(
                source_path,
                template_path,
                output_path,
            )

            generated_path = result.output_path
            if (
                not result.succeeded
                or generated_path is None
                or not generated_path.is_file()
            ):
                details = "\n".join(f"• {error}" for error in result.errors)
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
                messagebox.showerror("Error de generación", error_message)
                self.status_label.config(
                    text=(
                        "No se generó un archivo nuevo. "
                        f"Semanas omitidas: {result.omitted_weeks}."
                    ),
                    fg="red",
                )
                return

            self.status_label.config(
                text=f"Documento generado exitosamente: {generated_path.name}",
                fg="green",
            )

            messagebox.showinfo(
                "Éxito",
                f"Documento generado correctamente:\n\n"
                f"📄 {generated_path.name}\n"
                f"📁 {generated_path.parent}\n\n"
                f"Se escribieron {result.written_weeks} semanas.\n"
                f"Se omitieron {result.omitted_weeks} semanas.\n"
                f"Consulte 'meeting_generator.log' para más detalles.",
            )

            logger.info(
                "Proceso completado exitosamente. Archivo: %s",
                generated_path,
            )

        except Exception as exc:
            error_message = (
                f"Error al generar el documento:\n\n{str(exc)}\n\n"
                f"Consulte 'meeting_generator.log' para más detalles."
            )
            logger.error("Error en la generación: %s", exc, exc_info=True)
            messagebox.showerror("Error", error_message)
            self.status_label.config(
                text=f"Error: {str(exc)[:80]}",
                fg="red",
            )
        finally:
            self._set_generation_in_progress(False)

    def run(self) -> None:
        """Ejecuta la aplicación."""
        self.root.mainloop()


def main() -> None:
    """Función principal que inicia la aplicación."""
    # Configurar logging
    setup_logging()

    logger.info("=" * 60)
    logger.info("Meeting Generator - Iniciando aplicación")
    logger.info("=" * 60)

    # Crear y ejecutar la aplicación
    app = MeetingGeneratorApp()
    app.run()


if __name__ == "__main__":
    main()
