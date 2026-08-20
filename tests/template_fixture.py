"""Constructor de la plantilla DOCX sintética usada por las pruebas.

El documento reproduce únicamente el contrato estructural que consume
``TemplateWriter``. No deriva contenido, estilos ni metadatos de la plantilla
operativa y puede regenerarse en cualquier checkout.
"""

from pathlib import Path

from docx import Document
from docx.table import Table

FORM_COUNT = 10
_COLUMN_COUNT = 14
_TABLE_LAYOUTS = (
    (103, (2, 27, 54, 79), (0, 52)),
    (51, (2, 27), (0,)),
    (51, (2, 27), (0,)),
    (51, (2, 27), (0,)),
)


def build_template(path: Path) -> Path:
    """Genera una plantilla S-140 mínima con capacidad para diez semanas."""
    document = Document()
    document.core_properties.title = "Plantilla sintética para pruebas"
    document.core_properties.author = ""

    for row_count, form_starts, banner_rows in _TABLE_LAYOUTS:
        table = document.add_table(rows=row_count, cols=_COLUMN_COUNT)
        for row_index in banner_rows:
            _set_text(table, row_index, 0, "[NOMBRE DE LA CONGREGACIÓN]")
            _set_text(
                table,
                row_index,
                3,
                "Programa sintético para la reunión de entre semana",
            )
        for start_row in form_starts:
            _populate_form(table, start_row)

    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(path))
    return path


def _populate_form(table: Table, start_row: int) -> None:
    """Añade los marcadores de un formulario conservando los offsets reales."""
    _set_text(
        table,
        start_row,
        0,
        "[FECHA] | LECTURA SEMANAL DE LA BIBLIA",
    )
    _set_text(table, start_row, 7, "Presidente:")
    _set_text(table, start_row, 13, "[Nombre]")

    _set_text(table, start_row + 3, 0, "0:00")
    _set_text(table, start_row + 3, 2, "Canción [Número]")
    _set_text(table, start_row + 3, 9, "Oración:")
    _set_text(table, start_row + 3, 13, "[Nombre]")

    _set_text(table, start_row + 4, 0, "0:00")
    _set_text(table, start_row + 4, 2, "Palabras de introducción (1 min.)")

    _set_text(table, start_row + 6, 0, "TESOROS DE LA BIBLIA")
    _populate_assignment_row(
        table,
        start_row + 7,
        "1. [Título] (10 mins.)",
        "[Nombre]",
        title_column=2,
    )
    _populate_assignment_row(
        table,
        start_row + 8,
        "2. [Título] (10 mins.)",
        "[Nombre]",
        title_column=2,
    )
    _populate_assignment_row(
        table,
        start_row + 9,
        "3. [Título] (4 mins.)",
        "[Nombre]",
        title_column=2,
    )

    _set_text(table, start_row + 11, 0, "SEAMOS MEJORES MAESTROS")
    for offset, number in enumerate(range(4, 8), start=12):
        row_index = start_row + offset
        _populate_assignment_row(
            table,
            row_index,
            f"{number}. [Título] (X mins.)",
            "[Nombre/Nombre]",
            title_column=1,
        )
        _set_text(table, row_index, 10, "Estudiante/Ayudante:")

    _set_text(table, start_row + 17, 0, "NUESTRA VIDA CRISTIANA")
    _set_text(table, start_row + 18, 0, "0:00")
    _set_text(table, start_row + 18, 2, "Canción [Número]")

    for offset, number in ((19, 8), (20, 9)):
        _populate_assignment_row(
            table,
            start_row + offset,
            f"{number}. [Título] (XX mins.)",
            "[Nombre]",
            title_column=2,
        )

    bible_study_row = start_row + 21
    _populate_assignment_row(
        table,
        bible_study_row,
        "10. Estudio bíblico de la congregación (30 mins.)",
        "[Nombre/Nombre]",
        title_column=2,
    )
    _set_text(table, bible_study_row, 6, "Conductor/Lector:")

    _set_text(table, start_row + 22, 0, "0:00")
    _set_text(table, start_row + 22, 2, "Palabras de conclusión (3 mins.)")

    _set_text(table, start_row + 23, 0, "0:00")
    _set_text(table, start_row + 23, 2, "Canción [Número]")
    _set_text(table, start_row + 23, 9, "Oración:")
    _set_text(table, start_row + 23, 13, "[Nombre]")


def _populate_assignment_row(
    table: Table,
    row_index: int,
    assignment_text: str,
    participant_marker: str,
    *,
    title_column: int,
) -> None:
    _set_text(table, row_index, 0, "0:00")
    _set_text(table, row_index, title_column, assignment_text)
    _set_text(table, row_index, 13, participant_marker)


def _set_text(table: Table, row_index: int, column_index: int, text: str) -> None:
    table.rows[row_index].cells[column_index].text = text
