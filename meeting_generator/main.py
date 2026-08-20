"""
Punto de entrada del Meeting Generator.

Proporciona una interfaz gráfica con tkinter para seleccionar
el documento fuente y la plantilla, ejecuta el parser y el writer,
y genera el archivo S-140_COMPLETADO.docx.
"""

import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

from .config import OUTPUT_FILENAME, ALLOWED_EXTENSIONS
from .utils import setup_logging, get_default_output_path
from .parser import parse_document
from .template_writer import fill_template

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
        source_button = tk.Button(
            file_frame,
            text="1. Seleccionar Documento del Programa",
            command=self._select_source_file,
            width=40,
            height=2,
            bg="#e0e0e0",
        )
        source_button.pack(pady=5)

        self.source_label = tk.Label(
            file_frame,
            text="No se ha seleccionado archivo",
            font=("Arial", 9),
            fg="gray",
        )
        self.source_label.pack(pady=2)

        # Botón para seleccionar plantilla
        template_button = tk.Button(
            file_frame,
            text="2. Seleccionar Plantilla S-140",
            command=self._select_template_file,
            width=40,
            height=2,
            bg="#e0e0e0",
        )
        template_button.pack(pady=5)

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
            text="Generar S-140_COMPLETADO.docx",
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
                "El archivo generado se guardará como:\n"
                "S-140_COMPLETADO.docx en el mismo directorio de la plantilla."
            ),
            font=("Arial", 8),
            fg="gray",
        )
        output_info.pack()

    def _select_source_file(self) -> None:
        """Abre el diálogo para seleccionar el documento fuente del programa."""
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
        if self.source_path and self.template_path:
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

    def _generate_document(self) -> None:
        """
        Ejecuta el proceso de generación del documento completo.

        Realiza el parsing del documento fuente, rellena la plantilla
        y guarda el archivo de salida.
        """
        if not self.source_path or not self.template_path:
            messagebox.showerror(
                "Error",
                "Debe seleccionar ambos archivos antes de generar.",
            )
            return

        try:
            self.status_label.config(
                text="Analizando documento fuente...",
                fg="orange",
            )
            self.root.update()

            # Paso 1: Parsear el documento fuente
            logger.info("Iniciando procesamiento del documento fuente...")
            weeks = parse_document(self.source_path)

            if not weeks:
                messagebox.showwarning(
                    "Advertencia",
                    "No se encontraron semanas en el documento fuente.\n"
                    "Verifique que el documento tenga el formato esperado.",
                )
                self.status_label.config(
                    text="No se encontraron semanas. Verifique el documento.",
                    fg="red",
                )
                return

            self.status_label.config(
                text=f"Se encontraron {len(weeks)} semanas. Rellenando plantilla...",
                fg="orange",
            )
            self.root.update()

            # Paso 2: Determinar la ruta de salida
            output_path = get_default_output_path(
                self.source_path, self.template_path
            )

            # Paso 3: Rellenar la plantilla
            logger.info("Rellenando plantilla con %d semanas...", len(weeks))
            generated_path = fill_template(
                self.template_path,
                weeks,
                output_path,
            )

            # Paso 4: Notificar éxito
            self.status_label.config(
                text=f"Documento generado exitosamente: {generated_path.name}",
                fg="green",
            )

            messagebox.showinfo(
                "Éxito",
                f"Documento generado correctamente:\n\n"
                f"📄 {generated_path.name}\n"
                f"📁 {generated_path.parent}\n\n"
                f"Se procesaron {len(weeks)} semanas.\n"
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