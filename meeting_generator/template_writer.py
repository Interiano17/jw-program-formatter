"""
template_writer.py — Escritor profesional del documento S-140.

Arquitectura:
1. DESCUBRIMIENTO: escanea la plantilla buscando encabezados de sección
   ("SEAMOS MEJORES MAESTROS", "NUESTRA VIDA CRISTIANA") para localizar
   las regiones dinámicas sin depender de índices fijos.
2. ESCRITURA POR FILA: cada Assignment se escribe como una unidad completa
   en su fila correspondiente (número, título, duración, participantes).
3. LIMPIEZA: las filas de plantilla sin asignación correspondiente se limpian
   de placeholders.
4. VALIDACIÓN FINAL: barrido de todo el documento para garantizar 0 placeholders.

Columnas: 0-6 = [FECHA], 7-13 = resto. Sala auxiliar (8-12) se deja vacía.
Auditorio principal = columna 13.
"""

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from docx import Document

from .config import (
    CONGREGATION_NAME,
    DEFAULT_DURATIONS,
    MEETING_START_MINUTES,
    MINISTRY_TRANSITION_MINUTES,
    OPENING_PRAYER_DURATION_MINUTES,
    SONG_DURATION_MINUTES,
)
from .models import MeetingWeek, Assignment

logger = logging.getLogger(__name__)

PLACEHOLDERS = [
    "[FECHA]", "[Nombre]", "[Título]", "[Nombre/Nombre]",
    "Canción [Número]", "[Número]", "[X mins.]", "[XX mins.]",
    "(X mins.)", "(XX mins.)",
    "[NOMBRE DE LA CONGREGACIÓN]",
]

HEADER_TREASURES = "TESOROS DE LA BIBLIA"
HEADER_MINISTRY = "SEAMOS MEJORES MAESTROS"
HEADER_CHRISTIAN_LIFE = "NUESTRA VIDA CRISTIANA"
HEADER_CONCLUSION = "Palabras de conclusión"


@dataclass
class RowDescriptor:
    row_index: int
    number_cell: Optional[int] = None
    title_cell: Optional[int] = None
    duration_cell: Optional[int] = None
    participant_cell: Optional[int] = None
    auxiliary_cells: list[int] = field(default_factory=list)

    def content_cells(self) -> list[int]:
        cells: list[int] = []
        for cell_index in (
            self.number_cell,
            self.title_cell,
            self.duration_cell,
            self.participant_cell,
        ):
            if cell_index is not None and cell_index not in cells:
                cells.append(cell_index)
        return cells


class TemplateWriter:
    """Escritor profesional del documento S-140."""

    def __init__(self, template_path: Path):
        self.template_path = template_path
        self.document: Optional[Document] = None

    # ==================================================================
    # API pública
    # ==================================================================

    def fill(self, weeks: list[MeetingWeek], output_path: Path) -> Path:
        logger.info("Cargando plantilla: %s", self.template_path)
        self.document = Document(self.template_path)

        if not self.document.tables:
            logger.error("La plantilla no contiene tablas.")
            return output_path

        total_forms = sum(len(self._find_copy_starts(table)) for table in self.document.tables)
        logger.info("Capacidad detectada en la plantilla: %d formularios.", total_forms)

        weeks_to_fill = weeks[:total_forms]
        if len(weeks) > total_forms:
            logger.warning(
                "Solo %d formularios disponibles. Se usan %d de %d semanas.",
                total_forms,
                total_forms,
                len(weeks),
            )

        week_offset = 0
        for table_index, table in enumerate(self.document.tables, start=1):
            consumed = self._process_table(table, weeks_to_fill, week_offset, table_index)
            if consumed > 0:
                logger.info(
                    "Tabla %d: semanas índice %d a %d",
                    table_index,
                    week_offset,
                    week_offset + consumed - 1,
                )
            else:
                logger.debug("Tabla %d: no recibió semanas", table_index)
            week_offset += consumed

        self._write_congregation_name()
        self._clear_all_placeholders()
        self._clear_unused_time_markers()
        remaining = self._count_placeholders()
        logger.info("Placeholders tras limpieza: %d", remaining)

        self.document.save(str(output_path))
        logger.info("Documento guardado: %s", output_path)
        return output_path

    # ==================================================================
    # Procesamiento de una tabla
    # ==================================================================

    def _process_table(self, table, weeks: list[MeetingWeek], week_offset: int,
                       table_index: Optional[int] = None) -> int:
        """Procesa una tabla completa y devuelve cuántas semanas consumió."""
        copy_starts = self._find_copy_starts(table)
        if not copy_starts:
            return 0

        consumed = 0
        for ci, start_row in enumerate(copy_starts):
            week_index = week_offset + ci
            if week_index >= len(weeks):
                break
            end_row = copy_starts[ci + 1] if ci + 1 < len(copy_starts) else len(table.rows)
            self._process_one_copy(table, start_row, end_row, weeks[week_index])
            consumed += 1

        if table_index is not None and consumed == 0:
            logger.debug("Tabla %d: copy_starts=%d pero sin semanas disponibles", table_index, len(copy_starts))
        elif table_index is not None:
            logger.debug(
                "Tabla %d: consumió %d semanas (índices globales %d-%d)",
                table_index,
                consumed,
                week_offset,
                week_offset + consumed - 1,
            )

        return consumed

    def _find_copy_starts(self, table) -> list[int]:
        starts = []
        for ri, row in enumerate(table.rows):
            if len(row.cells) > 0 and "[FECHA]" in row.cells[0].text:
                starts.append(ri)
        return starts

    # ==================================================================
    # Procesamiento de UNA copia del formulario
    # ==================================================================

    def _process_one_copy(self, table, start_row: int, end_row: int, week: MeetingWeek) -> None:
        """Rellena una copia del formulario."""
        structure = self._discover_structure(table, start_row, end_row)
        if not structure:
            logger.warning("Estructura no descubierta en filas %d-%d", start_row, end_row)
            return

        # 1. Cabecera fija
        self._write_header(table, week, structure)

        # 2. Tesoros de la Biblia
        self._write_section_rows(table, week.bible_treasures, structure["bible_treasures_rows"])

        # 3. Ministerio — escribir los existentes, limpiar sobrantes
        self._write_section_rows(table, week.ministry, structure["ministry_rows"])

        # 4. Vida Cristiana (solo normales, no BS ni conclusión)
        cv_normal = [a for a in week.christian_life
                     if not a.is_bible_study and not a.is_conclusion]
        self._write_section_rows(table, cv_normal, structure["cv_normal_rows"])

        # 5. Estudio bíblico
        if structure.get("bible_study_row") is not None:
            eb = next((a for a in week.christian_life if a.is_bible_study), None)
            if eb:
                self._write_bible_study_row(table, structure["bible_study_row"], eb)
            else:
                self._clear_row(
                    table,
                    structure["bible_study_row"].row_index,
                )

        # 6. Conclusión
        if structure.get("conclusion_row") is not None:
            self._clear_row(table, structure["conclusion_row"])

        # 7. Canciones y oraciones
        self._write_songs_and_prayers(table, week, structure)

        # 8. Horas de inicio
        self._write_schedule_times(table, start_row, end_row, week, structure)

    # ==================================================================
    # Descubrimiento de estructura
    # ==================================================================

    def _discover_structure(self, table, start_row: int, end_row: int) -> Optional[dict]:
        """Escanea las filas y descubre la posición de cada elemento."""
        struct = {
            "fecha_row": None,
            "presidente_name_row": None,
            "consejero_name_row": None,
            "cancion_inicio_row": None,
            "oracion_inicio_name_row": None,
            "introduction_row": None,
            "bible_treasures_rows": [],
            "ministry_rows": [],
            "cv_normal_rows": [],
            "bible_study_row": None,
            "conclusion_row": None,
            "cancion_intermedia_row": None,
            "cancion_final_row": None,
            "oracion_final_name_row": None,
        }

        current_section = "header"

        for ri in range(start_row, end_row):
            row = table.rows[ri]
            combined = " | ".join(
                row.cells[ci].text.strip()
                for ci in range(min(14, len(row.cells)))
            )

            if HEADER_MINISTRY in combined.upper() and current_section != "christian_life":
                current_section = "ministry"
                continue
            if HEADER_CHRISTIAN_LIFE in combined.upper():
                current_section = "christian_life"
                continue

            if current_section == "header":
                if "[FECHA]" in combined:
                    struct["fecha_row"] = ri
                if "Presidente:" in combined:
                    struct["presidente_name_row"] = self._find_nombre_row(table, ri, end_row)
                if "Consejero" in combined and struct["consejero_name_row"] is None:
                    for ci in range(7, 14):
                        if ci < len(row.cells) and "[Nombre]" in row.cells[ci].text:
                            struct["consejero_name_row"] = ri
                            break
                if "Canción [Número]" in combined and struct["cancion_inicio_row"] is None:
                    struct["cancion_inicio_row"] = ri
                if "Oración:" in combined and struct["oracion_inicio_name_row"] is None:
                    struct["oracion_inicio_name_row"] = self._find_nombre_in_row(row, ri)
                if "Palabras de introducción" in combined:
                    struct["introduction_row"] = ri
                if (
                    ("[Título]" in combined or "[Nombre/Nombre]" in combined or "[Nombre]" in combined)
                    and "Presidente:" not in combined
                    and "Consejero" not in combined
                    and "Oración:" not in combined
                    and "Canción [Número]" not in combined
                ):
                    struct["bible_treasures_rows"].append(self._build_row_descriptor(table, ri))

            elif current_section == "ministry":
                if "[Título]" in combined or "[Nombre/Nombre]" in combined:
                    struct["ministry_rows"].append(self._build_row_descriptor(table, ri))

            elif current_section == "christian_life":
                if "Canción [Número]" in combined and struct["cancion_intermedia_row"] is None:
                    struct["cancion_intermedia_row"] = ri
                    continue
                if "Estudio bíblico" in combined:
                    struct["bible_study_row"] = self._build_row_descriptor(table, ri)
                    continue
                if "[Título]" in combined or "[Nombre]" in combined:
                    struct["cv_normal_rows"].append(self._build_row_descriptor(table, ri))
                    continue
                if HEADER_CONCLUSION in combined:
                    struct["conclusion_row"] = ri
                    current_section = "closing"
                    continue

            elif current_section == "closing":
                if "Canción [Número]" in combined and struct["cancion_final_row"] is None:
                    struct["cancion_final_row"] = ri
                if "Oración:" in combined and struct["oracion_final_name_row"] is None:
                    struct["oracion_final_name_row"] = self._find_nombre_in_row(row, ri)

        return struct

    def _build_row_descriptor(self, table, row_idx: int) -> RowDescriptor:
        row = table.rows[row_idx]
        descriptor = RowDescriptor(row_index=row_idx)
        participant_cells_by_identity: dict[int, int] = {}

        for ci, cell in enumerate(row.cells):
            text = cell.text.strip()
            if not text:
                continue

            if text == "Sala auxiliar":
                descriptor.auxiliary_cells.append(ci)

            if "[Nombre/Nombre]" in text or "[Nombre]" in text:
                # Una celda física puede repetirse en row.cells por gridSpan; conservamos
                # la última posición vista de cada tc para escoger luego la de mayor índice.
                participant_cells_by_identity[id(cell._tc)] = ci

            if descriptor.title_cell is None and (
                "[Título]" in text or "Estudio bíblico" in text
            ):
                descriptor.title_cell = ci
                descriptor.duration_cell = ci
                if descriptor.number_cell is None and self._starts_with_number(text):
                    descriptor.number_cell = ci
                continue

            if descriptor.number_cell is None and self._starts_with_number(text):
                descriptor.number_cell = ci

            if descriptor.duration_cell is None and self._has_duration_marker(text):
                descriptor.duration_cell = ci

        if descriptor.title_cell is None:
            descriptor.title_cell = descriptor.number_cell
        if descriptor.duration_cell is None:
            descriptor.duration_cell = descriptor.title_cell
        if descriptor.number_cell is None:
            descriptor.number_cell = descriptor.title_cell

        if participant_cells_by_identity:
            participant_indices = sorted(participant_cells_by_identity.values())
            descriptor.participant_cell = participant_indices[-1]
            descriptor.auxiliary_cells.extend(
                ci for ci in participant_indices[:-1] if ci not in descriptor.auxiliary_cells
            )

        logger.debug(
            "Descriptor row=%s number_cell=%s title_cell=%s duration_cell=%s participant_cell=%s auxiliary_cells=%s",
            descriptor.row_index,
            descriptor.number_cell,
            descriptor.title_cell,
            descriptor.duration_cell,
            descriptor.participant_cell,
            descriptor.auxiliary_cells,
        )
        return descriptor

    def _starts_with_number(self, text: str) -> bool:
        stripped = text.lstrip()
        digits: list[str] = []
        for char in stripped:
            if char.isdigit():
                digits.append(char)
                continue
            break
        if not digits:
            return False
        remaining = stripped[len(digits):].lstrip()
        return remaining.startswith(".")

    def _has_duration_marker(self, text: str) -> bool:
        return any(marker in text for marker in ("[X mins.]", "[XX mins.]", "(X mins.)", "(XX mins.)"))

    def _find_nombre_row(self, table, ri: int, end_row: int) -> Optional[int]:
        """Busca la fila con [Nombre] a partir de ri (puede ser la misma o siguiente)."""
        for offset in range(2):
            check = ri + offset
            if check < end_row and check < len(table.rows):
                for ci in range(7, 14):
                    if ci < len(table.rows[check].cells):
                        if "[Nombre]" in table.rows[check].cells[ci].text:
                            return check
        return None

    def _find_nombre_in_row(self, row, row_idx: int) -> Optional[int]:
        """Busca [Nombre] en la misma fila, retorna el índice de fila o None."""
        for ci in range(7, 14):
            if ci < len(row.cells) and "[Nombre]" in row.cells[ci].text:
                return row_idx
        return None

    # ==================================================================
    # Escritura de cabecera fija
    # ==================================================================

    def _write_header(self, table, week: MeetingWeek, struct: dict) -> None:
        """Escribe los campos fijos de la cabecera."""
        # [FECHA]
        if struct["fecha_row"] is not None and week.date:
            self._set_cell_text(table, struct["fecha_row"], 0, week.date, "[FECHA]")

        # Lectura semanal
        if struct["fecha_row"] is not None and week.weekly_reading:
            self._set_cell_text(
                table,
                struct["fecha_row"],
                0,
                week.weekly_reading,
                "LECTURA SEMANAL DE LA BIBLIA",
            )

        # Presidente
        if struct["presidente_name_row"] is not None and week.president:
            self._write_nombre_in_row(table, struct["presidente_name_row"], week.president)

        # Consejero
        if struct["consejero_name_row"] is not None:
            self._write_nombre_in_row(table, struct["consejero_name_row"], week.president)

        # Canción inicial
        if struct["cancion_inicio_row"] is not None and week.opening_song:
            self._write_cancion_in_row(table, struct["cancion_inicio_row"], week.opening_song)

        # Oración inicial
        if struct["oracion_inicio_name_row"] is not None:
            name = week.opening_prayer or week.president
            self._write_nombre_in_row(table, struct["oracion_inicio_name_row"], name)

    # ==================================================================
    # Escritura de secciones dinámicas
    # ==================================================================

    def _write_section_rows(self, table, assignments: list[Assignment],
                              row_descriptors: list[RowDescriptor]) -> None:
        """Escribe asignaciones en las filas existentes; limpia las sobrantes."""
        ordered_assignments = sorted(assignments, key=lambda assignment: assignment.number)
        for i, descriptor in enumerate(row_descriptors):
            if descriptor.row_index >= len(table.rows):
                continue
            if i < len(ordered_assignments):
                self._write_assignment_in_row(table, descriptor, ordered_assignments[i])
            else:
                self._clear_row(table, descriptor.row_index)
                self._clear_time_in_row(table, descriptor.row_index)

    def _write_assignment_in_row(self, table, descriptor: RowDescriptor, assignment: Assignment) -> None:
        """Escribe una asignación completa en una fila."""
        row = table.rows[descriptor.row_index]

        self._clear_auxiliary_cells(row, descriptor)

        participant_cell = descriptor.participant_cell

        for ci, cell in enumerate(row.cells):
            text = cell.text
            if "[Número]" in text or self._starts_with_number(text):
                self._write_number_in_cell(cell, assignment.number)

            if "[Título]" in text:
                self._write_title_in_cell(cell, assignment.title)

            if self._has_duration_marker(text):
                self._write_duration_in_cell(cell, assignment.duration_text)

            if "[Nombre/Nombre]" in text or "[Nombre]" in text:
                if participant_cell is not None and ci == participant_cell:
                    # La celda principal es la de mayor índice entre las que contienen
                    # el placeholder de participante; las anteriores se vacían.
                    logger.debug(
                        "Fila %s: participantes escritos en columna %s",
                        descriptor.row_index,
                        ci,
                    )
                    self._write_participants_in_cell(cell, assignment)
                else:
                    logger.debug(
                        "Fila %s: celda auxiliar vaciada en columna %s",
                        descriptor.row_index,
                        ci,
                    )
                    self._empty_cell(cell)

    def _write_bible_study_row(self, table, descriptor: RowDescriptor, assignment: Assignment) -> None:
        """Escribe el estudio bíblico."""
        self._write_assignment_in_row(table, descriptor, assignment)

    def _clear_auxiliary_cells(self, row, descriptor: RowDescriptor) -> None:
        seen_cells: set[int] = set()
        for ci in descriptor.auxiliary_cells:
            if ci < len(row.cells):
                cell = row.cells[ci]
                cell_id = id(cell._tc)
                if cell_id in seen_cells:
                    continue
                seen_cells.add(cell_id)
                self._empty_cell(cell)

    def _empty_cell(self, cell) -> None:
        for par in cell.paragraphs:
            for run in par.runs:
                run.text = ""

    def _write_number_in_cell(self, cell, number: int) -> None:
        if self._replace_in_cell_conditional(cell, "[Número]", str(number)):
            return
        for par in cell.paragraphs:
            full = self._paragraph_text(par)
            prefix = self._leading_number_prefix(full)
            if prefix:
                self._replace_in_paragraph(par, prefix, f"{number}.")
                return

    def _write_title_in_cell(self, cell, title: str) -> None:
        if not title:
            return
        if self._replace_in_cell_conditional(cell, "[Título]", title):
            return
        for par in cell.paragraphs:
            full = self._paragraph_text(par)
            if title in full:
                return
            prefix = self._leading_number_prefix(full)
            if not prefix:
                continue
            duration_index = self._first_duration_marker_index(full)
            if duration_index == -1:
                continue
            middle = full[len(prefix):duration_index]
            search = f"{prefix}{middle}"
            replacement = f"{prefix} {title} "
            if self._replace_in_paragraph(par, search, replacement):
                return

    def _write_duration_in_cell(self, cell, duration_text: str) -> None:
        if not duration_text:
            return
        for placeholder in ("[X mins.]", "[XX mins.]", "(X mins.)", "(XX mins.)"):
            replacement = duration_text if placeholder.startswith("[") else f"({duration_text})"
            if self._replace_in_cell_conditional(cell, placeholder, replacement):
                return

    def _write_participants_in_cell(self, cell, assignment: Assignment) -> None:
        participants = assignment.formatted_participants()
        self._replace_in_cell_conditional(cell, "[Nombre/Nombre]", participants)
        self._replace_in_cell_conditional(cell, "[Nombre]", participants)

    # ==================================================================
    # Canciones y oraciones
    # ==================================================================

    def _write_songs_and_prayers(self, table, week: MeetingWeek, struct: dict) -> None:
        # Canción intermedia
        if struct["cancion_intermedia_row"] is not None and week.intermediate_song:
            self._write_cancion_in_row(table, struct["cancion_intermedia_row"],
                                        week.intermediate_song)
        # Canción final
        if struct["cancion_final_row"] is not None and week.closing_song:
            self._write_cancion_in_row(table, struct["cancion_final_row"], week.closing_song)
        # Oración final
        if struct["oracion_final_name_row"] is not None and week.closing_prayer:
            self._write_nombre_in_row(table, struct["oracion_final_name_row"],
                                       week.closing_prayer)

    # ==================================================================
    # Helpers de escritura en fila
    # ==================================================================

    def _write_nombre_in_row(self, table, row_idx: int, name: str) -> None:
        if row_idx >= len(table.rows):
            return
        for ci in range(7, 14):
            if ci < len(table.rows[row_idx].cells):
                cell = table.rows[row_idx].cells[ci]
                if "[Nombre]" in cell.text and "[Nombre/Nombre]" not in cell.text:
                    self._replace_in_cell(cell, "[Nombre]", name)
                    return

    def _write_cancion_in_row(self, table, row_idx: int, song: str) -> None:
        if row_idx >= len(table.rows):
            return
        seen_cells: set[int] = set()
        for cell in table.rows[row_idx].cells:
            cell_id = id(cell._tc)
            if cell_id in seen_cells:
                continue
            seen_cells.add(cell_id)
            if "Canción [Número]" in cell.text:
                self._replace_in_cell(cell, "Canción [Número]", f"Canción {song}")
                return

    # ==================================================================
    # Cálculo de horas
    # ==================================================================

    def _write_schedule_times(
        self,
        table,
        start_row: int,
        end_row: int,
        week: MeetingWeek,
        struct: dict,
    ) -> None:
        """Calcula y escribe la hora de inicio de cada actividad."""
        events: list[tuple[int, int]] = []

        self._add_schedule_event(
            events,
            struct.get("cancion_inicio_row"),
            SONG_DURATION_MINUTES + OPENING_PRAYER_DURATION_MINUTES,
        )
        self._add_schedule_event(
            events,
            struct.get("introduction_row"),
            DEFAULT_DURATIONS["introduccion"],
        )
        self._add_assignment_events(
            events,
            week.bible_treasures,
            struct["bible_treasures_rows"],
        )
        self._add_assignment_events(
            events,
            week.ministry,
            struct["ministry_rows"],
            additional_minutes=MINISTRY_TRANSITION_MINUTES,
        )
        self._add_schedule_event(
            events,
            struct.get("cancion_intermedia_row"),
            SONG_DURATION_MINUTES,
        )

        christian_life = [
            assignment
            for assignment in week.christian_life
            if not assignment.is_bible_study and not assignment.is_conclusion
        ]
        self._add_assignment_events(
            events,
            christian_life,
            struct["cv_normal_rows"],
        )

        bible_study = next(
            (
                assignment
                for assignment in week.christian_life
                if assignment.is_bible_study
            ),
            None,
        )
        if bible_study is not None:
            self._add_schedule_event(
                events,
                self._descriptor_row(struct.get("bible_study_row")),
                self._assignment_duration(bible_study),
            )

        self._add_schedule_event(
            events,
            struct.get("conclusion_row"),
            DEFAULT_DURATIONS["conclusion"],
        )
        self._add_schedule_event(
            events,
            struct.get("cancion_final_row"),
            SONG_DURATION_MINUTES,
        )

        events.sort(key=lambda event: event[0])
        active_rows = {row_index for row_index, _ in events}
        for row_index in range(start_row, min(end_row, len(table.rows))):
            if row_index not in active_rows:
                self._clear_time_in_row(table, row_index)

        current_minutes = MEETING_START_MINUTES
        for row_index, duration_minutes in events:
            self._write_time_in_row(table, row_index, current_minutes)
            current_minutes += duration_minutes

    def _add_assignment_events(
        self,
        events: list[tuple[int, int]],
        assignments: list[Assignment],
        descriptors: list[RowDescriptor],
        additional_minutes: int = 0,
    ) -> None:
        ordered_assignments = sorted(assignments, key=lambda assignment: assignment.number)
        for assignment, descriptor in zip(ordered_assignments, descriptors):
            duration = self._assignment_duration(assignment) + additional_minutes
            self._add_schedule_event(events, descriptor.row_index, duration)

    def _assignment_duration(self, assignment: Assignment) -> int:
        if assignment.duration_minutes > 0:
            return assignment.duration_minutes
        logger.warning(
            "Punto %d sin duración; no avanzará el cálculo de hora.",
            assignment.number,
        )
        return 0

    def _add_schedule_event(
        self,
        events: list[tuple[int, int]],
        row_index: Optional[int],
        duration_minutes: int,
    ) -> None:
        if row_index is not None:
            events.append((row_index, duration_minutes))

    def _descriptor_row(self, descriptor: Optional[RowDescriptor]) -> Optional[int]:
        return descriptor.row_index if descriptor is not None else None

    def _write_time_in_row(self, table, row_idx: int, total_minutes: int) -> None:
        if row_idx >= len(table.rows) or not table.rows[row_idx].cells:
            return
        cell = table.rows[row_idx].cells[0]
        current_text = cell.text.strip()
        display_text = self._format_time(total_minutes)
        if current_text:
            self._replace_in_cell(cell, current_text, display_text)

    def _clear_time_in_row(self, table, row_idx: int) -> None:
        if row_idx >= len(table.rows) or not table.rows[row_idx].cells:
            return
        cell = table.rows[row_idx].cells[0]
        if cell.text.strip() == "0:00":
            self._remove_text(cell, "0:00")

    def _format_time(self, total_minutes: int) -> str:
        hours, minutes = divmod(total_minutes, 60)
        display_hour = hours % 12 or 12
        return f"{display_hour}:{minutes:02d}"

    # ==================================================================
    # Limpieza de filas y celdas
    # ==================================================================

    def _clear_row(self, table, row_idx: int) -> None:
        if row_idx >= len(table.rows):
            return
        for cell in table.rows[row_idx].cells:
            self._clear_cell(cell)

    def _clear_cell(self, cell) -> None:
        for ph in PLACEHOLDERS:
            if ph in cell.text:
                self._remove_text(cell, ph)

    def _clear_all_placeholders(self) -> None:
        if self.document is None:
            return
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for ph in PLACEHOLDERS:
                        if ph in cell.text:
                            self._remove_text(cell, ph)

    def _write_congregation_name(self) -> None:
        if self.document is None:
            return
        marker = "[NOMBRE DE LA CONGREGACIÓN]"
        for table in self.document.tables:
            seen_cells: set[int] = set()
            for row in table.rows:
                for cell in row.cells:
                    cell_id = id(cell._tc)
                    if cell_id in seen_cells:
                        continue
                    seen_cells.add(cell_id)
                    if marker in cell.text:
                        self._replace_in_cell(cell, marker, CONGREGATION_NAME)

    def _clear_unused_time_markers(self) -> None:
        if self.document is None:
            return
        for table in self.document.tables:
            for row_index, _row in enumerate(table.rows):
                self._clear_time_in_row(table, row_index)

    def _count_placeholders(self) -> int:
        if self.document is None:
            return 0
        count = 0
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for ph in PLACEHOLDERS:
                        if ph in cell.text:
                            count += 1
        return count

    # ==================================================================
    # Reemplazo de texto preservando formato (nivel de runs)
    # ==================================================================

    def _set_cell_text(self, table, row_idx: int, col_idx: int, new_text: str,
                        old_text: str = "") -> None:
        if row_idx >= len(table.rows) or col_idx >= len(table.rows[row_idx].cells):
            return
        cell = table.rows[row_idx].cells[col_idx]
        search = old_text if old_text else cell.text.strip()
        if search and search in cell.text:
            self._replace_in_cell(cell, search, new_text)

    def _replace_in_cell_conditional(self, cell, old: str, new: str) -> bool:
        """Reemplaza solo si old está presente y new no está vacío."""
        if not new or old not in cell.text:
            return False
        return self._replace_in_cell(cell, old, new)

    def _paragraph_text(self, par) -> str:
        return "".join(run.text for run in par.runs)

    def _leading_number_prefix(self, text: str) -> str:
        stripped = text.lstrip()
        digits: list[str] = []
        for char in stripped:
            if char.isdigit():
                digits.append(char)
                continue
            break
        if not digits:
            return ""
        remaining = stripped[len(digits):].lstrip()
        if not remaining.startswith("."):
            return ""
        return f"{''.join(digits)}."

    def _first_duration_marker_index(self, text: str) -> int:
        markers = ("[X mins.]", "[XX mins.]", "(X mins.)", "(XX mins.)")
        positions = [text.find(marker) for marker in markers if text.find(marker) != -1]
        if not positions:
            return -1
        return min(positions)

    def _replace_in_cell(self, cell, old: str, new: str) -> bool:
        if not new:
            return False
        for par in cell.paragraphs:
            if self._replace_in_paragraph(par, old, new):
                return True
        return False

    def _remove_text(self, cell, text: str) -> None:
        for par in cell.paragraphs:
            self._replace_in_paragraph(par, text, "")

    def _replace_in_paragraph(self, par, search: str, repl: str) -> bool:
        runs = par.runs
        if not runs:
            return False
        full, b = "", []
        for ri, r in enumerate(runs):
            if r.text:
                s = len(full)
                full += r.text
                b.append((ri, s, len(full)))
        pos = full.find(search)
        if pos == -1:
            return False
        end = pos + len(search)
        aff = []
        for ri, rs, re in b:
            if rs <= pos < re:
                aff.append(ri)
            elif pos <= rs < end:
                aff.append(ri)
            elif rs < end <= re and ri not in aff:
                aff.append(ri)
        if not aff:
            return False
        aff.sort()
        f = aff[0]
        fs = next(s for r, s, e in b if r == f)
        rel = pos - fs
        orig = runs[f].text
        runs[f].text = orig[:rel] + repl + orig[rel + len(search):]
        for ri in aff[1:]:
            rs = next(s for r, s, e in b if r == ri)
            re = next(e for r, s, e in b if r == ri)
            if rs >= pos and re <= end:
                runs[ri].text = ""
            elif rs < end <= re:
                ov = end - rs
                if 0 < ov <= len(runs[ri].text):
                    runs[ri].text = runs[ri].text[ov:]
            elif rs <= pos < re:
                ov = re - pos
                if 0 < ov <= len(runs[ri].text):
                    runs[ri].text = runs[ri].text[:-ov]
        return True


def fill_template(template_path: Path, weeks: list[MeetingWeek], output_path: Path) -> Path:
    writer = TemplateWriter(template_path)
    return writer.fill(weeks, output_path)
