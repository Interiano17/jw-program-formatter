"""
template_writer.py — Escritor del documento S-140.

Arquitectura:
1. DESCUBRIMIENTO Y PREFLIGHT: localiza las copias y valida su estructura y
   capacidad antes de modificarlas.
2. ESCRITURA POR FILA: cada Assignment se escribe como una unidad completa
   en su fila correspondiente (número, título, duración, participantes).
3. VERIFICACIÓN: comprueba valores escritos y marcadores en cada copia usada.
4. LIMPIEZA FINAL: elimina formularios y filas sobrantes, incluidas las tablas
   cuyos formularios no se usan, y guarda el documento libre de marcadores.

Columnas: 0-6 = [FECHA], 7-13 = resto. Sala auxiliar (8-12) se deja vacía.
Auditorio principal = columna 13.
"""

import hashlib
import logging
import os
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, TypedDict

from docx import Document
from docx.document import Document as DocxDocument
from docx.table import Table, _Cell, _Row
from docx.text.paragraph import Paragraph

from .config import (
    CONGREGATION_NAME,
    DEFAULT_DURATIONS,
    INTERMEDIATE_SONG_DURATION_OVERRIDES,
    MEETING_START_MINUTES,
    MINISTRY_TRANSITION_MINUTES,
    OPENING_PRAYER_DURATION_MINUTES,
    PLACEHOLDERS,
    SONG_DURATION_MINUTES,
    TEMPLATE_MARKERS,
)
from .models import Assignment, MeetingWeek
from .results import GenerationResult
from .validation import GenerationValidationError

logger = logging.getLogger(__name__)

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
    writes_title: bool = False
    writes_duration: bool = False
    auxiliary_cells: list[int] = field(default_factory=list)


class TemplateStructure(TypedDict):
    fecha_row: Optional[int]
    presidente_name_row: Optional[int]
    consejero_name_row: Optional[int]
    cancion_inicio_row: Optional[int]
    oracion_inicio_name_row: Optional[int]
    introduction_row: Optional[int]
    bible_treasures_rows: list[RowDescriptor]
    ministry_rows: list[RowDescriptor]
    cv_normal_rows: list[RowDescriptor]
    bible_study_row: Optional[RowDescriptor]
    conclusion_row: Optional[RowDescriptor]
    cancion_intermedia_row: Optional[int]
    cancion_final_row: Optional[int]
    oracion_final_name_row: Optional[int]


@dataclass(frozen=True)
class TemplateCopy:
    table: Table
    table_index: int
    start_row: int
    end_row: int
    structure: TemplateStructure


class DocumentPublicationError(OSError):
    """Conserva la ruta publicada si falla una operación posterior al reemplazo."""

    def __init__(self, output_path: Path, cause: Exception) -> None:
        self.output_path = output_path
        super().__init__(
            f"El documento se publicó, pero falló una operación posterior: {cause}"
        )


class TemplateWriter:
    """Escritor del documento S-140."""

    def __init__(self, template_path: Path) -> None:
        self.template_path = template_path
        self.document: Optional[DocxDocument] = None

    # ==================================================================
    # API pública
    # ==================================================================

    def fill(
        self,
        weeks: list[MeetingWeek],
        output_path: Path,
    ) -> GenerationResult:
        """Ejecuta la generación completa y devuelve un resultado verificable."""
        try:
            if self.template_path.resolve() == output_path.resolve():
                raise GenerationValidationError(
                    "Ruta de salida inválida",
                    ["La ruta de salida coincide con la plantilla"],
                )
            created_path = self._fill_document(weeks, output_path)
        except DocumentPublicationError as exc:
            errors = [str(exc)]
            try:
                logger.error(
                    "Falló una operación posterior a la publicación", exc_info=True
                )
            except Exception as log_error:
                errors.append(
                    f"No se pudo registrar el error de publicación: {log_error}"
                )
            return GenerationResult.published_failure(
                len(weeks), exc.output_path, errors
            )
        except GenerationValidationError as exc:
            logger.error("Generación rechazada: %s", exc)
            return GenerationResult.failure(len(weeks), list(exc.issues))
        except Exception as exc:
            logger.error("La generación falló", exc_info=True)
            return GenerationResult.failure(
                len(weeks),
                [f"No se pudo generar el documento: {exc}"],
            )

        return GenerationResult.success(len(weeks), created_path)

    def _fill_document(self, weeks: list[MeetingWeek], output_path: Path) -> Path:
        logger.info("Cargando plantilla.")
        self.document = Document(str(self.template_path))

        copies = self._discover_template_copies()
        logger.info("Capacidad detectada en la plantilla: %d formularios.", len(copies))
        issues = self._input_validation_issues(weeks, copies)
        if issues:
            raise GenerationValidationError("No se puede generar el documento", issues)

        write_issues: list[str] = []
        schedule_rows: list[tuple[_Row, str, str]] = []
        for week, copy in zip(weeks, copies):
            self._process_one_copy(
                copy.table,
                copy.start_row,
                copy.end_row,
                week,
                copy.structure,
            )
            write_issues.extend(self._written_content_issues(week, copy))
            rows = list(copy.table.rows)
            for row_index, total_minutes in self._schedule_start_times(
                week, copy.structure
            ):
                schedule_rows.append(
                    (
                        rows[row_index],
                        self._format_time(total_minutes),
                        f"semana {week.week_index}, tabla {copy.table_index}, "
                        f"fila {row_index + 1}",
                    )
                )

        remaining = self._clear_unused_copies_and_find_markers(copies, len(weeks))
        if remaining:
            write_issues.append(f"sin resolver: {', '.join(remaining)}")
        if write_issues:
            raise GenerationValidationError(
                "La plantilla no se pudo completar", write_issues
            )
        self._remove_unused_structure(copies, weeks)
        # Las filas conservan su identidad, pero sus índices cambian al limpiar.
        row_locations = {
            row._tr: (table_index, row_index)
            for table_index, table in enumerate(self.document.tables)
            for row_index, row in enumerate(table.rows)
        }
        expected_times: list[tuple[int, int, str, str]] = []
        for row, expected_time, context in schedule_rows:
            location = row_locations.get(row._tr)
            if location is None:
                write_issues.append(f"{context}: se eliminó una actividad del horario")
            else:
                expected_times.append((*location, expected_time, context))
        write_issues.extend(self._final_schedule_issues(self.document, expected_times))
        if write_issues:
            raise GenerationValidationError(
                "El horario final no se pudo completar", write_issues
            )
        logger.info("Marcadores residuales tras la validación: 0")

        return self._save_document_atomically(output_path, expected_times)

    def _save_document_atomically(
        self, output_path: Path, expected_times: list[tuple[int, int, str, str]]
    ) -> Path:
        """Publica un DOCX verificado sin dejar una salida parcial."""
        if self.document is None:
            raise GenerationValidationError(
                "No se puede guardar el documento",
                ["no hay un documento preparado para guardar"],
            )
        if not output_path.parent.is_dir():
            raise GenerationValidationError(
                "No se puede guardar el documento",
                [f"el directorio de salida no existe: {output_path.parent}"],
            )

        created_path = output_path.parent.resolve(strict=True) / output_path.name
        descriptor, temporary_name = tempfile.mkstemp(
            dir=str(created_path.parent),
            prefix=f".{output_path.stem}-",
            suffix=".docx",
        )
        os.close(descriptor)
        temporary_path = Path(temporary_name)
        published = False
        try:
            self.document.save(str(temporary_path))
            if not temporary_path.is_file() or temporary_path.stat().st_size <= 0:
                raise OSError("el archivo temporal no se creó correctamente")
            saved_document = Document(str(temporary_path))
            schedule_issues = self._final_schedule_issues(
                saved_document, expected_times
            )
            if schedule_issues:
                raise GenerationValidationError(
                    "El documento guardado contiene un horario inválido",
                    schedule_issues,
                )
            expected_digest = self._file_digest(temporary_path)
            os.replace(temporary_path, created_path)
            published = True
            if not created_path.is_file() or created_path.stat().st_size <= 0:
                raise OSError("no se pudo verificar el archivo de salida")
            if self._file_digest(created_path) != expected_digest:
                raise OSError("el archivo publicado no coincide con el generado")
            logger.info("Documento guardado y verificado.")
            return created_path
        except Exception as exc:
            if published:
                raise DocumentPublicationError(created_path, exc) from exc
            raise
        finally:
            if not published:
                temporary_path.unlink(missing_ok=True)

    def _file_digest(self, path: Path) -> bytes:
        digest = hashlib.sha256()
        with path.open("rb") as document_file:
            for chunk in iter(lambda: document_file.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.digest()

    # ==================================================================
    # Descubrimiento y validación de formularios
    # ==================================================================

    def _discover_template_copies(self) -> list[TemplateCopy]:
        if self.document is None:
            return []

        copies: list[TemplateCopy] = []
        for table_index, table in enumerate(self.document.tables, start=1):
            rows = list(table.rows)
            starts = self._find_copy_starts(rows)
            for copy_index, start_row in enumerate(starts):
                end_row = (
                    starts[copy_index + 1]
                    if copy_index + 1 < len(starts)
                    else len(rows)
                )
                copies.append(
                    TemplateCopy(
                        table=table,
                        table_index=table_index,
                        start_row=start_row,
                        end_row=end_row,
                        structure=self._discover_structure(table, start_row, end_row),
                    )
                )
        return copies

    def _input_validation_issues(
        self,
        weeks: list[MeetingWeek],
        copies: list[TemplateCopy],
    ) -> list[str]:
        issues: list[str] = []
        if not weeks:
            issues.append("no hay semanas para escribir")
        if not copies:
            issues.append("la plantilla no contiene formularios reconocibles")
        if len(weeks) > len(copies):
            issues.append(
                f"la capacidad de la plantilla es de {len(copies)} semanas, "
                f"pero se recibieron {len(weeks)}"
            )

        for week in weeks:
            if not week.is_valid():
                issues.extend(
                    f"semana {week.week_index}: {issue}"
                    for issue in week.validation_errors()
                )

        for week, copy in zip(weeks, copies):
            issues.extend(self._copy_validation_issues(week, copy))
        return issues

    def _copy_validation_issues(
        self,
        week: MeetingWeek,
        copy: TemplateCopy,
    ) -> list[str]:
        prefix = f"semana {week.week_index}, tabla {copy.table_index}"
        structure = copy.structure
        rows = list(copy.table.rows)
        issues = self._copy_structure_issues(prefix, copy, rows)

        christian_life_normal = self._normal_christian_life_assignments(week)
        assignment_rows: list[tuple[Assignment, RowDescriptor]] = []
        capacities = (
            (
                "Tesoros de la Biblia",
                week.bible_treasures,
                structure["bible_treasures_rows"],
            ),
            ("Seamos Mejores Maestros", week.ministry, structure["ministry_rows"]),
            (
                "Nuestra Vida Cristiana",
                christian_life_normal,
                structure["cv_normal_rows"],
            ),
        )
        for section_name, assignments, descriptors in capacities:
            if len(assignments) > len(descriptors):
                issues.append(
                    f"{prefix}: la capacidad de {section_name} es de "
                    f"{len(descriptors)} filas y se necesitan {len(assignments)}"
                )
            for assignment, descriptor in zip(assignments, descriptors):
                assignment_rows.append((assignment, descriptor))
                issues.extend(
                    self._descriptor_validation_issues(
                        prefix,
                        section_name,
                        copy.table,
                        descriptor,
                        assignment,
                    )
                )

        bible_study_row = structure["bible_study_row"]
        if bible_study_row is not None:
            bible_study = next(
                (
                    assignment
                    for assignment in week.christian_life
                    if assignment.is_bible_study
                ),
                None,
            )
        else:
            bible_study = None
        if bible_study_row is not None and bible_study is not None:
            assignment_rows.append((bible_study, bible_study_row))
            issues.extend(
                self._descriptor_validation_issues(
                    prefix,
                    "Estudio bíblico",
                    copy.table,
                    bible_study_row,
                    bible_study,
                )
            )

        conclusion_row = structure["conclusion_row"]
        conclusions = [
            assignment for assignment in week.christian_life if assignment.is_conclusion
        ]
        if conclusion_row is not None:
            template_duration = self._fixed_duration_in_row(
                copy.table,
                conclusion_row.row_index,
            )
            expected_duration = DEFAULT_DURATIONS["conclusion"]
            if template_duration != expected_duration:
                issues.append(
                    f"{prefix}: la conclusión de la plantilla debe durar "
                    f"{expected_duration} minutos"
                )
            if (
                len(conclusions) == 1
                and template_duration is not None
                and conclusions[0].duration_minutes != template_duration
            ):
                issues.append(
                    f"{prefix}: la duración de la conclusión no coincide con la "
                    "plantilla"
                )
            if conclusion_row.participant_cell is None:
                issues.append(
                    f"{prefix}, fila {conclusion_row.row_index + 1}: la conclusión "
                    "requiere un marcador de participante o una celda final vacía "
                    "en el lado derecho"
                )
            if len(conclusions) == 1:
                assignment_rows.append((conclusions[0], conclusion_row))

        issues.extend(self._shared_participant_issues(prefix, rows, assignment_rows))
        return issues

    def _copy_structure_issues(
        self, prefix: str, copy: TemplateCopy, rows: list[_Row]
    ) -> list[str]:
        """Valida las filas y rótulos requeridos del formulario."""
        issues: list[str] = []
        structure = copy.structure
        required_rows = (
            ("fecha", structure["fecha_row"]),
            ("presidente", structure["presidente_name_row"]),
            ("canción inicial", structure["cancion_inicio_row"]),
            ("oración inicial", structure["oracion_inicio_name_row"]),
            ("introducción", structure["introduction_row"]),
            ("canción intermedia", structure["cancion_intermedia_row"]),
            ("estudio bíblico", structure["bible_study_row"]),
            ("conclusión", structure["conclusion_row"]),
            ("canción final", structure["cancion_final_row"]),
            ("oración final", structure["oracion_final_name_row"]),
        )
        for field_name, row in required_rows:
            if row is None:
                issues.append(f"{prefix}: falta la fila de {field_name}")

        required_labels = (
            "LECTURA SEMANAL DE LA BIBLIA",
            HEADER_TREASURES,
            HEADER_MINISTRY,
            HEADER_CHRISTIAN_LIFE,
            HEADER_CONCLUSION,
        )
        row_texts = [
            text
            for row in rows[copy.start_row : copy.end_row]
            for text in self._unique_cell_texts(row)
        ]
        for label in required_labels:
            if not any(label in text for text in row_texts):
                issues.append(f"{prefix}: falta el marcador '{label}'")
        return issues

    def _shared_participant_issues(
        self,
        prefix: str,
        rows: list[_Row],
        assignment_rows: list[tuple[Assignment, RowDescriptor]],
    ) -> list[str]:
        """Rechaza participantes incompatibles que comparten una celda física."""
        issues: list[str] = []
        participant_values: dict[object, str] = {}
        for assignment, descriptor in assignment_rows:
            if descriptor.participant_cell is None:
                continue
            cells = rows[descriptor.row_index].cells
            cell = cells[descriptor.participant_cell]
            participants = assignment.formatted_participants()
            previous = participant_values.setdefault(cell._tc, participants)
            if previous != participants:
                label = (
                    "la conclusión"
                    if assignment.is_conclusion
                    else f"el punto {assignment.number}"
                )
                issues.append(
                    f"{prefix}, fila {descriptor.row_index + 1}: la celda fusionada "
                    f"requiere participantes distintos para {label}"
                )
        return issues

    def _normal_christian_life_assignments(self, week: MeetingWeek) -> list[Assignment]:
        return [
            assignment
            for assignment in week.christian_life
            if not assignment.is_bible_study and not assignment.is_conclusion
        ]

    def _descriptor_validation_issues(
        self,
        prefix: str,
        section_name: str,
        table: Table,
        descriptor: RowDescriptor,
        assignment: Assignment,
    ) -> list[str]:
        row = table.rows[descriptor.row_index]
        texts = self._unique_cell_texts(row)
        missing: list[str] = []
        if not any(self._starts_with_number(text) for text in texts):
            missing.append("número")
        if not any(self._has_writable_title(text) for text in texts):
            missing.append("título")
        if not any(
            self._has_duration_marker(text)
            or self._fixed_duration_from_text(text) is not None
            for text in texts
        ):
            missing.append("duración")
        if descriptor.participant_cell is None or not any(
            "[Nombre]" in text or "[Nombre/Nombre]" in text for text in texts
        ):
            missing.append("participantes")

        issues = []
        if missing:
            issues.append(
                f"{prefix}: una fila de {section_name} no admite {', '.join(missing)}"
            )

        fixed_duration = next(
            (
                duration
                for text in texts
                if (duration := self._fixed_duration_from_text(text)) is not None
            ),
            None,
        )
        if fixed_duration is not None and assignment.duration_minutes != fixed_duration:
            issues.append(
                f"{prefix}: el punto {assignment.number} dura "
                f"{assignment.duration_minutes} minutos, pero su fila fija "
                f"muestra {fixed_duration}"
            )
        return issues

    def _unique_cell_texts(self, row: _Row) -> list[str]:
        texts: list[str] = []
        seen_cells: set[object] = set()
        for cell in row.cells:
            if cell._tc in seen_cells:
                continue
            seen_cells.add(cell._tc)
            texts.append(cell.text.strip())
        return texts

    def _physical_row_cells(
        self, row: _Row, cells: tuple[_Cell, ...]
    ) -> list[tuple[int, _Cell]]:
        """Devuelve una vez las celdas cuyo contenido pertenece a esta fila."""
        physical_cells: list[tuple[int, _Cell]] = []
        seen_cells: set[object] = set()
        for cell_index, cell in enumerate(cells):
            # Las continuaciones verticales apuntan al tc de la fila de origen.
            if cell._tc in seen_cells or cell._tc.getparent() is not row._tr:
                continue
            seen_cells.add(cell._tc)
            physical_cells.append((cell_index, cell))
        return physical_cells

    def _has_writable_title(self, text: str) -> bool:
        if "[Título]" in text:
            return True
        prefix = self._leading_number_prefix(text)
        if not prefix:
            return False
        duration_match = re.search(r"\(\s*\d+\s+mins?\.?\s*\)", text, re.IGNORECASE)
        if duration_match is None:
            return False
        return bool(text[len(prefix) : duration_match.start()].strip())

    def _fixed_duration_from_text(self, text: str) -> Optional[int]:
        match = re.search(r"\(\s*(\d+)\s+mins?\.?\s*\)", text, re.IGNORECASE)
        return int(match.group(1)) if match else None

    def _fixed_duration_in_row(self, table: Table, row_index: int) -> Optional[int]:
        for text in self._unique_cell_texts(table.rows[row_index]):
            duration = self._fixed_duration_from_text(text)
            if duration is not None:
                return duration
        return None

    def _written_content_issues(
        self,
        week: MeetingWeek,
        copy: TemplateCopy,
    ) -> list[str]:
        """Comprueba que los datos esperados estén presentes tras escribir."""
        issues: list[str] = []
        structure = copy.structure
        prefix = f"semana {week.week_index}, tabla {copy.table_index}"

        row_values = (
            ("fecha", structure["fecha_row"], week.date),
            ("lectura semanal", structure["fecha_row"], week.weekly_reading),
            ("presidente", structure["presidente_name_row"], week.president),
            (
                "oración inicial",
                structure["oracion_inicio_name_row"],
                week.opening_prayer,
            ),
            (
                "oración final",
                structure["oracion_final_name_row"],
                week.closing_prayer,
            ),
        )
        for field_name, row_index, expected_value in row_values:
            if row_index is not None and not self._row_contains(
                copy.table,
                row_index,
                expected_value,
            ):
                issues.append(f"{prefix}: no se escribió {field_name}")

        songs = (
            ("canción inicial", structure["cancion_inicio_row"], week.opening_song),
            (
                "canción intermedia",
                structure["cancion_intermedia_row"],
                week.intermediate_song,
            ),
            ("canción final", structure["cancion_final_row"], week.closing_song),
        )
        for field_name, row_index, song_number in songs:
            if row_index is not None and not self._row_contains_pattern(
                copy.table,
                row_index,
                rf"Canción\s+{re.escape(song_number)}(?!\d)",
            ):
                issues.append(f"{prefix}: no se escribió {field_name}")

        christian_life_normal = self._normal_christian_life_assignments(week)
        sections = (
            (week.bible_treasures, structure["bible_treasures_rows"]),
            (week.ministry, structure["ministry_rows"]),
            (christian_life_normal, structure["cv_normal_rows"]),
        )
        for assignments, descriptors in sections:
            for assignment, descriptor in zip(
                sorted(assignments, key=lambda item: item.number),
                descriptors,
            ):
                issues.extend(
                    self._assignment_written_issues(
                        prefix,
                        copy.table,
                        descriptor,
                        assignment,
                    )
                )

        bible_study = next(
            (
                assignment
                for assignment in week.christian_life
                if assignment.is_bible_study
            ),
            None,
        )
        bible_study_row = structure["bible_study_row"]
        if bible_study is not None and bible_study_row is not None:
            issues.extend(
                self._assignment_written_issues(
                    prefix,
                    copy.table,
                    bible_study_row,
                    bible_study,
                )
            )
        conclusion = next(
            (
                assignment
                for assignment in week.christian_life
                if assignment.is_conclusion
            ),
            None,
        )
        conclusion_row = structure["conclusion_row"]
        if conclusion is not None and conclusion_row is not None:
            issues.extend(
                self._assignment_written_issues(
                    prefix, copy.table, conclusion_row, conclusion
                )
            )
        issues.extend(self._schedule_written_issues(week, copy))
        return issues

    def _assignment_written_issues(
        self,
        prefix: str,
        table: Table,
        descriptor: RowDescriptor,
        assignment: Assignment,
    ) -> list[str]:
        issues: list[str] = []
        row = table.rows[descriptor.row_index]
        cells = row.cells
        label = (
            "la conclusión"
            if assignment.is_conclusion
            else f"el punto {assignment.number}"
        )

        number_cell = descriptor.number_cell
        if number_cell is not None and not self._cell_starts_with_number(
            cells[number_cell],
            assignment.number,
        ):
            issues.append(f"{prefix}: no se escribió el número de {label}")

        title_cell = descriptor.title_cell
        if (
            descriptor.writes_title
            and title_cell is not None
            and assignment.title not in cells[title_cell].text
        ):
            issues.append(f"{prefix}: no se escribió el título de {label}")

        duration_cell = descriptor.duration_cell
        if (
            descriptor.writes_duration
            and duration_cell is not None
            and assignment.duration_text not in cells[duration_cell].text
        ):
            issues.append(f"{prefix}: no se escribió la duración de {label}")

        participant_cell = descriptor.participant_cell
        participants = assignment.formatted_participants()
        if (
            participant_cell is not None
            and participants not in cells[participant_cell].text
        ):
            issues.append(f"{prefix}: no se escribieron los participantes de {label}")
        return issues

    def _cell_starts_with_number(self, cell: _Cell, expected_number: int) -> bool:
        expected_prefix = f"{expected_number}."
        return any(
            self._leading_number_prefix(paragraph.text) == expected_prefix
            for paragraph in cell.paragraphs
        )

    def _row_contains(self, table: Table, row_index: int, text: str) -> bool:
        return any(
            text in cell_text
            for cell_text in self._unique_cell_texts(table.rows[row_index])
        )

    def _row_contains_pattern(
        self,
        table: Table,
        row_index: int,
        pattern: str,
    ) -> bool:
        return any(
            re.search(pattern, cell_text) is not None
            for cell_text in self._unique_cell_texts(table.rows[row_index])
        )

    def _find_copy_starts(self, rows: list[_Row]) -> list[int]:
        starts = []
        for row_index, row in enumerate(rows):
            cells = row.cells
            if (
                cells
                and cells[0]._tc.getparent() is row._tr
                and "[FECHA]" in cells[0].text
            ):
                starts.append(row_index)
        return starts

    # ==================================================================
    # Procesamiento de un formulario
    # ==================================================================

    def _process_one_copy(
        self,
        table: Table,
        start_row: int,
        end_row: int,
        week: MeetingWeek,
        structure: TemplateStructure,
    ) -> None:
        """Rellena una copia del formulario."""
        self._write_header(table, week, structure)

        self._write_section_rows(
            table, week.bible_treasures, structure["bible_treasures_rows"]
        )

        self._write_section_rows(table, week.ministry, structure["ministry_rows"])

        christian_life_normal = self._normal_christian_life_assignments(week)
        self._write_section_rows(
            table, christian_life_normal, structure["cv_normal_rows"]
        )

        bible_study_row = structure["bible_study_row"]
        if bible_study_row is not None:
            bible_study = next(
                (
                    assignment
                    for assignment in week.christian_life
                    if assignment.is_bible_study
                ),
                None,
            )
            if bible_study:
                self._write_bible_study_row(table, bible_study_row, bible_study)
            else:
                self._clear_row(
                    table,
                    bible_study_row.row_index,
                )

        conclusion = next(
            (
                assignment
                for assignment in week.christian_life
                if assignment.is_conclusion
            ),
            None,
        )
        conclusion_row = structure["conclusion_row"]
        if conclusion is not None and conclusion_row is not None:
            self._write_assignment_in_row(table, conclusion_row, conclusion)

        self._write_songs_and_prayers(table, week, structure)

        self._write_schedule_times(table, start_row, end_row, week, structure)

    # ==================================================================
    # Descubrimiento de estructura
    # ==================================================================

    def _discover_structure(
        self, table: Table, start_row: int, end_row: int
    ) -> TemplateStructure:
        """Escanea las filas y descubre la posición de cada elemento."""
        structure: TemplateStructure = {
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
        rows = list(table.rows)

        for row_index, row in enumerate(rows[start_row:end_row], start=start_row):
            cells = row.cells
            physical_cells = self._physical_row_cells(row, cells)
            row_text = " | ".join(
                cell.text.strip()
                for cell_index, cell in physical_cells
                if cell_index < 14
            )

            if (
                HEADER_MINISTRY in row_text.upper()
                and current_section != "christian_life"
            ):
                current_section = "ministry"
                continue
            if HEADER_CHRISTIAN_LIFE in row_text.upper():
                current_section = "christian_life"
                continue

            if current_section == "header":
                if "[FECHA]" in row_text:
                    structure["fecha_row"] = row_index
                if "Presidente:" in row_text:
                    structure["presidente_name_row"] = self._find_nombre_row(
                        rows, row_index, end_row
                    )
                if "Consejero" in row_text and structure["consejero_name_row"] is None:
                    for cell in cells[7:14]:
                        if "[Nombre]" in cell.text:
                            structure["consejero_name_row"] = row_index
                            break
                if (
                    "Canción [Número]" in row_text
                    and structure["cancion_inicio_row"] is None
                ):
                    structure["cancion_inicio_row"] = row_index
                if (
                    "Oración:" in row_text
                    and structure["oracion_inicio_name_row"] is None
                ):
                    structure["oracion_inicio_name_row"] = self._find_nombre_in_row(
                        cells, row_index
                    )
                if "Palabras de introducción" in row_text:
                    structure["introduction_row"] = row_index
                if (
                    (
                        "[Título]" in row_text
                        or "[Nombre/Nombre]" in row_text
                        or "[Nombre]" in row_text
                    )
                    and "Presidente:" not in row_text
                    and "Consejero" not in row_text
                    and "Oración:" not in row_text
                    and "Canción [Número]" not in row_text
                ):
                    structure["bible_treasures_rows"].append(
                        self._build_row_descriptor(cells, row_index)
                    )

            elif current_section == "ministry":
                if "[Título]" in row_text or "[Nombre/Nombre]" in row_text:
                    structure["ministry_rows"].append(
                        self._build_row_descriptor(cells, row_index)
                    )

            elif current_section == "christian_life":
                if (
                    "Canción [Número]" in row_text
                    and structure["cancion_intermedia_row"] is None
                ):
                    structure["cancion_intermedia_row"] = row_index
                    continue
                if "Estudio bíblico" in row_text:
                    structure["bible_study_row"] = self._build_row_descriptor(
                        cells, row_index
                    )
                    continue
                if HEADER_CONCLUSION in row_text:
                    descriptor = self._build_row_descriptor(cells, row_index)
                    if descriptor.participant_cell is None:
                        if (
                            physical_cells
                            and physical_cells[-1][0] >= 7
                            and not physical_cells[-1][1].text.strip()
                        ):
                            descriptor.participant_cell = physical_cells[-1][0]
                    structure["conclusion_row"] = descriptor
                    current_section = "closing"
                    continue
                if "[Título]" in row_text or "[Nombre]" in row_text:
                    structure["cv_normal_rows"].append(
                        self._build_row_descriptor(cells, row_index)
                    )
                    continue

            elif current_section == "closing":
                if (
                    "Canción [Número]" in row_text
                    and structure["cancion_final_row"] is None
                ):
                    structure["cancion_final_row"] = row_index
                if (
                    "Oración:" in row_text
                    and structure["oracion_final_name_row"] is None
                ):
                    structure["oracion_final_name_row"] = self._find_nombre_in_row(
                        cells, row_index
                    )

        return structure

    def _build_row_descriptor(
        self, cells: tuple[_Cell, ...], row_index: int
    ) -> RowDescriptor:
        descriptor = RowDescriptor(row_index=row_index)
        participant_indices: list[int] = []
        seen_cells: set[object] = set()

        for cell_index, cell in enumerate(cells):
            if cell._tc in seen_cells:
                continue
            seen_cells.add(cell._tc)
            text = cell.text.strip()
            if not text:
                continue

            if text == "Sala auxiliar":
                descriptor.auxiliary_cells.append(cell_index)

            if "[Nombre/Nombre]" in text or "[Nombre]" in text:
                participant_indices.append(cell_index)

            if descriptor.title_cell is None and (
                "[Título]" in text or "Estudio bíblico" in text
            ):
                descriptor.title_cell = cell_index
                descriptor.duration_cell = cell_index
                descriptor.writes_title = "[Título]" in text
                descriptor.writes_duration = self._has_duration_marker(text)
                if descriptor.number_cell is None and self._starts_with_number(text):
                    descriptor.number_cell = cell_index
                continue

            if descriptor.number_cell is None and self._starts_with_number(text):
                descriptor.number_cell = cell_index

            if descriptor.duration_cell is None and self._has_duration_marker(text):
                descriptor.duration_cell = cell_index

            if "[Título]" in text:
                descriptor.writes_title = True
            if self._has_duration_marker(text):
                descriptor.writes_duration = True

        if descriptor.title_cell is None:
            descriptor.title_cell = descriptor.number_cell
        if descriptor.duration_cell is None:
            descriptor.duration_cell = descriptor.title_cell
        if descriptor.number_cell is None:
            descriptor.number_cell = descriptor.title_cell

        if participant_indices:
            descriptor.participant_cell = participant_indices[-1]
            descriptor.auxiliary_cells.extend(
                cell_index
                for cell_index in participant_indices[:-1]
                if cell_index not in descriptor.auxiliary_cells
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
        for character in stripped:
            if character.isdigit():
                digits.append(character)
                continue
            break
        if not digits:
            return False
        remaining = stripped[len(digits) :].lstrip()
        return remaining.startswith(".")

    def _has_duration_marker(self, text: str) -> bool:
        return any(
            marker in text
            for marker in ("[X mins.]", "[XX mins.]", "(X mins.)", "(XX mins.)")
        )

    def _find_nombre_row(
        self, rows: list[_Row], row_index: int, end_row: int
    ) -> Optional[int]:
        """Busca [Nombre] en la fila indicada o en la siguiente."""
        for offset in range(2):
            candidate_row_index = row_index + offset
            if candidate_row_index < end_row and candidate_row_index < len(rows):
                cells = rows[candidate_row_index].cells
                if any("[Nombre]" in cell.text for cell in cells[7:14]):
                    return candidate_row_index
        return None

    def _find_nombre_in_row(
        self, cells: tuple[_Cell, ...], row_index: int
    ) -> Optional[int]:
        """Busca [Nombre] en la misma fila, retorna el índice de fila o None."""
        if any("[Nombre]" in cell.text for cell in cells[7:14]):
            return row_index
        return None

    # ==================================================================
    # Escritura de cabecera fija
    # ==================================================================

    def _write_header(
        self, table: Table, week: MeetingWeek, structure: TemplateStructure
    ) -> None:
        """Escribe los campos fijos de la cabecera."""
        if structure["fecha_row"] is not None and week.date:
            self._set_cell_text(table, structure["fecha_row"], 0, week.date, "[FECHA]")

        if structure["fecha_row"] is not None and week.weekly_reading:
            self._set_cell_text(
                table,
                structure["fecha_row"],
                0,
                week.weekly_reading,
                "LECTURA SEMANAL DE LA BIBLIA",
            )

        if structure["presidente_name_row"] is not None and week.president:
            self._write_nombre_in_row(
                table, structure["presidente_name_row"], week.president
            )

        if structure["consejero_name_row"] is not None:
            self._write_nombre_in_row(
                table, structure["consejero_name_row"], week.president
            )

        if structure["cancion_inicio_row"] is not None and week.opening_song:
            self._write_cancion_in_row(
                table, structure["cancion_inicio_row"], week.opening_song
            )

        if structure["oracion_inicio_name_row"] is not None:
            name = week.opening_prayer or week.president
            self._write_nombre_in_row(table, structure["oracion_inicio_name_row"], name)

    # ==================================================================
    # Escritura de secciones dinámicas
    # ==================================================================

    def _write_section_rows(
        self,
        table: Table,
        assignments: list[Assignment],
        row_descriptors: list[RowDescriptor],
    ) -> None:
        """Escribe asignaciones en las filas existentes; limpia las sobrantes."""
        ordered_assignments = sorted(
            assignments, key=lambda assignment: assignment.number
        )
        row_count = len(table.rows)
        for assignment_index, descriptor in enumerate(row_descriptors):
            if descriptor.row_index >= row_count:
                continue
            if assignment_index < len(ordered_assignments):
                self._write_assignment_in_row(
                    table, descriptor, ordered_assignments[assignment_index]
                )
            else:
                self._clear_row(table, descriptor.row_index)
                self._clear_time_in_row(table, descriptor.row_index)

    def _write_assignment_in_row(
        self, table: Table, descriptor: RowDescriptor, assignment: Assignment
    ) -> None:
        """Escribe una asignación completa en una fila."""
        row = table.rows[descriptor.row_index]
        cells = row.cells
        participant_tc = (
            cells[descriptor.participant_cell]._tc
            if descriptor.participant_cell is not None
            else None
        )
        auxiliary_tcs = {
            cells[cell_index]._tc
            for cell_index in descriptor.auxiliary_cells
            if cell_index < len(cells)
        }

        for cell_index, cell in self._physical_row_cells(row, cells):
            if cell._tc in auxiliary_tcs and cell._tc is not participant_tc:
                self._empty_cell(cell)
                continue
            text = cell.text
            if "[Número]" in text or self._starts_with_number(text):
                self._write_number_in_cell(cell, assignment.number)

            if "[Título]" in text:
                self._write_title_in_cell(cell, assignment.title)

            if self._has_duration_marker(text):
                self._write_duration_in_cell(cell, assignment.duration_text)

            if cell._tc is participant_tc:
                logger.debug(
                    "Fila %s: participantes escritos en columna %s",
                    descriptor.row_index,
                    cell_index,
                )
                self._write_participants_in_cell(cell, assignment)
            elif "[Nombre/Nombre]" in text or "[Nombre]" in text:
                logger.debug(
                    "Fila %s: celda auxiliar vaciada en columna %s",
                    descriptor.row_index,
                    cell_index,
                )
                self._empty_cell(cell)

    def _write_bible_study_row(
        self, table: Table, descriptor: RowDescriptor, assignment: Assignment
    ) -> None:
        """Escribe el estudio bíblico."""
        self._write_assignment_in_row(table, descriptor, assignment)

    def _empty_cell(self, cell: _Cell) -> None:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.text = ""

    def _write_number_in_cell(self, cell: _Cell, number: int) -> None:
        if self._replace_in_cell_conditional(cell, "[Número]", str(number)):
            return
        for paragraph in cell.paragraphs:
            paragraph_text = self._paragraph_text(paragraph)
            prefix = self._leading_number_prefix(paragraph_text)
            if prefix:
                self._replace_in_paragraph(paragraph, prefix, f"{number}.")
                return

    def _write_title_in_cell(self, cell: _Cell, title: str) -> None:
        if not title:
            return
        if self._replace_in_cell_conditional(cell, "[Título]", title):
            return
        for paragraph in cell.paragraphs:
            paragraph_text = self._paragraph_text(paragraph)
            if title in paragraph_text:
                return
            prefix = self._leading_number_prefix(paragraph_text)
            if not prefix:
                continue
            duration_index = self._first_duration_marker_index(paragraph_text)
            if duration_index == -1:
                continue
            middle = paragraph_text[len(prefix) : duration_index]
            search = f"{prefix}{middle}"
            replacement = f"{prefix} {title} "
            if self._replace_in_paragraph(paragraph, search, replacement):
                return

    def _write_duration_in_cell(self, cell: _Cell, duration_text: str) -> None:
        if not duration_text:
            return
        for placeholder in ("[X mins.]", "[XX mins.]", "(X mins.)", "(XX mins.)"):
            replacement = (
                duration_text if placeholder.startswith("[") else f"({duration_text})"
            )
            if self._replace_in_cell_conditional(cell, placeholder, replacement):
                return

    def _write_participants_in_cell(self, cell: _Cell, assignment: Assignment) -> None:
        participants = assignment.formatted_participants()
        if not cell.text.strip():
            paragraph = cell.paragraphs[0]
            run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            run.text = participants
            return
        self._replace_in_cell_conditional(cell, "[Nombre/Nombre]", participants)
        self._replace_in_cell_conditional(cell, "[Nombre]", participants)

    # ==================================================================
    # Canciones y oraciones
    # ==================================================================

    def _write_songs_and_prayers(
        self, table: Table, week: MeetingWeek, structure: TemplateStructure
    ) -> None:
        if structure["cancion_intermedia_row"] is not None and week.intermediate_song:
            self._write_cancion_in_row(
                table, structure["cancion_intermedia_row"], week.intermediate_song
            )
        if structure["cancion_final_row"] is not None and week.closing_song:
            self._write_cancion_in_row(
                table, structure["cancion_final_row"], week.closing_song
            )
        if structure["oracion_final_name_row"] is not None and week.closing_prayer:
            self._write_nombre_in_row(
                table, structure["oracion_final_name_row"], week.closing_prayer
            )

    # ==================================================================
    # Helpers de escritura en fila
    # ==================================================================

    def _write_nombre_in_row(self, table: Table, row_index: int, name: str) -> None:
        rows = list(table.rows)
        if row_index >= len(rows):
            return
        cells = rows[row_index].cells
        for cell in cells[7:14]:
            text = cell.text
            if "[Nombre]" in text and "[Nombre/Nombre]" not in text:
                self._replace_in_cell(cell, "[Nombre]", name)
                return

    def _write_cancion_in_row(self, table: Table, row_index: int, song: str) -> None:
        rows = list(table.rows)
        if row_index >= len(rows):
            return
        row = rows[row_index]
        cells = row.cells
        for _, cell in self._physical_row_cells(row, cells):
            if "Canción [Número]" in cell.text:
                self._replace_in_cell(cell, "Canción [Número]", f"Canción {song}")
                return

    # ==================================================================
    # Cálculo de horas
    # ==================================================================

    def _write_schedule_times(
        self,
        table: Table,
        start_row: int,
        end_row: int,
        week: MeetingWeek,
        structure: TemplateStructure,
    ) -> None:
        """Escribe las horas calculadas y limpia las filas sin actividad."""
        start_times = self._schedule_start_times(week, structure)
        active_rows = {row_index for row_index, _ in start_times}
        for row_index in range(start_row, min(end_row, len(table.rows))):
            if row_index not in active_rows:
                self._clear_time_in_row(table, row_index)

        for row_index, total_minutes in start_times:
            self._write_time_in_row(table, row_index, total_minutes)

    def _schedule_start_times(
        self, week: MeetingWeek, structure: TemplateStructure
    ) -> list[tuple[int, int]]:
        start_times: list[tuple[int, int]] = []
        current_minutes = MEETING_START_MINUTES
        for row_index, duration_minutes in self._schedule_events(week, structure):
            start_times.append((row_index, current_minutes))
            current_minutes += duration_minutes
        return start_times

    def _schedule_written_issues(
        self, week: MeetingWeek, copy: TemplateCopy
    ) -> list[str]:
        """Verifica cada hora y rechaza fusiones entre actividades incompatibles."""
        issues: list[str] = []
        rows = list(copy.table.rows)
        time_cells: dict[object, tuple[int, int]] = {}
        prefix = f"semana {week.week_index}, tabla {copy.table_index}"
        for row_index, total_minutes in self._schedule_start_times(
            week, copy.structure
        ):
            expected_time = self._format_time(total_minutes)
            cell = self._time_cell(rows[row_index])
            if cell is None:
                issues.append(f"{prefix}, fila {row_index + 1}: falta la celda horaria")
                continue
            previous = time_cells.get(cell._tc)
            if previous is not None:
                previous_row, previous_minutes = previous
                if previous_minutes != total_minutes:
                    issues.append(
                        f"{prefix}, fila {row_index + 1}: la celda horaria fusionada "
                        f"con la fila {previous_row + 1} requiere horas distintas "
                        f"({self._format_time(previous_minutes)} y {expected_time})"
                    )
                continue
            time_cells[cell._tc] = (row_index, total_minutes)
            if cell.text.strip() != expected_time:
                issues.append(
                    f"{prefix}, fila {row_index + 1}: no se escribió la hora "
                    f"esperada {expected_time}"
                )
        return issues

    def _final_schedule_issues(
        self,
        document: DocxDocument,
        expected_times: list[tuple[int, int, str, str]],
    ) -> list[str]:
        """Comprueba las horas tras limpiar filas y después de serializar el DOCX."""
        issues: list[str] = []
        tables_rows = [list(table.rows) for table in document.tables]
        for table_index, row_index, expected_time, context in expected_times:
            if table_index >= len(tables_rows) or row_index >= len(
                tables_rows[table_index]
            ):
                issues.append(f"{context}: falta la fila de una actividad del horario")
                continue
            cell = self._time_cell(tables_rows[table_index][row_index])
            if cell is None:
                issues.append(f"{context}: falta la celda horaria")
            elif cell.text.strip() != expected_time:
                issues.append(
                    f"{context}: no se escribió la hora esperada {expected_time}"
                )
        return issues

    def _schedule_events(
        self, week: MeetingWeek, structure: TemplateStructure
    ) -> list[tuple[int, int]]:
        """Calcula las filas y duraciones del horario sin modificar la plantilla."""
        events: list[tuple[int, int]] = []

        self._add_schedule_event(
            events,
            structure.get("cancion_inicio_row"),
            SONG_DURATION_MINUTES + OPENING_PRAYER_DURATION_MINUTES,
        )
        self._add_schedule_event(
            events,
            structure.get("introduction_row"),
            DEFAULT_DURATIONS["introduccion"],
        )
        self._add_assignment_events(
            events,
            week.bible_treasures,
            structure["bible_treasures_rows"],
        )
        self._add_assignment_events(
            events,
            week.ministry,
            structure["ministry_rows"],
            additional_minutes=MINISTRY_TRANSITION_MINUTES,
        )
        self._add_schedule_event(
            events,
            structure.get("cancion_intermedia_row"),
            INTERMEDIATE_SONG_DURATION_OVERRIDES.get(
                (week.date, week.weekly_reading), SONG_DURATION_MINUTES
            ),
        )

        christian_life = self._normal_christian_life_assignments(week)
        self._add_assignment_events(
            events,
            christian_life,
            structure["cv_normal_rows"],
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
                self._descriptor_row(structure.get("bible_study_row")),
                self._assignment_duration(bible_study),
            )

        self._add_schedule_event(
            events,
            self._descriptor_row(structure.get("conclusion_row")),
            DEFAULT_DURATIONS["conclusion"],
        )
        self._add_schedule_event(
            events,
            structure.get("cancion_final_row"),
            SONG_DURATION_MINUTES,
        )

        events.sort(key=lambda event: event[0])
        return events

    def _add_assignment_events(
        self,
        events: list[tuple[int, int]],
        assignments: list[Assignment],
        descriptors: list[RowDescriptor],
        additional_minutes: int = 0,
    ) -> None:
        ordered_assignments = sorted(
            assignments, key=lambda assignment: assignment.number
        )
        for assignment, descriptor in zip(ordered_assignments, descriptors):
            duration = self._assignment_duration(assignment) + additional_minutes
            self._add_schedule_event(events, descriptor.row_index, duration)

    def _assignment_duration(self, assignment: Assignment) -> int:
        if (
            isinstance(assignment.duration_minutes, int)
            and not isinstance(assignment.duration_minutes, bool)
            and assignment.duration_minutes > 0
        ):
            return assignment.duration_minutes
        raise GenerationValidationError(
            "No se puede calcular el horario",
            [f"el punto {assignment.number} no tiene una duración positiva"],
        )

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

    def _time_cell(self, row: _Row) -> Optional[_Cell]:
        """Devuelve la celda horaria, o None si la fila omite esa columna."""
        grid_before = row._tr.xpath("w:trPr/w:gridBefore/@w:val")
        if grid_before and int(grid_before[0]) > 0:
            return None
        cells = row.cells
        return cells[0] if cells else None

    def _write_time_in_row(
        self, table: Table, row_index: int, total_minutes: int
    ) -> None:
        rows = list(table.rows)
        if row_index >= len(rows):
            return
        row = rows[row_index]
        cell = self._time_cell(row)
        if cell is None:
            return
        current_text = cell.text.strip()
        if (
            current_text
            and re.fullmatch(r"(?:\d{1,2}:\d{2}\s*)+", current_text) is None
        ):
            return
        display_text = self._format_time(total_minutes)
        if current_text:
            self._replace_in_cell(cell, current_text, display_text)
        else:
            paragraph = cell.paragraphs[0]
            run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
            run.text = display_text

    def _clear_time_in_row(self, table: Table, row_index: int) -> None:
        rows = list(table.rows)
        if row_index >= len(rows):
            return
        row = rows[row_index]
        cell = self._time_cell(row)
        if cell is None:
            return
        if cell._tc.getparent() is not row._tr:
            return
        current_text = cell.text.strip()
        if re.fullmatch(r"\d{1,2}:\d{2}", current_text):
            self._remove_text(cell, current_text)

    def _format_time(self, total_minutes: int) -> str:
        hours, minutes = divmod(total_minutes, 60)
        display_hour = hours % 12 or 12
        return f"{display_hour}:{minutes:02d}"

    # ==================================================================
    # Limpieza de filas y celdas
    # ==================================================================

    def _clear_row(self, table: Table, row_index: int) -> None:
        rows = list(table.rows)
        if row_index >= len(rows):
            return
        row = rows[row_index]
        cells = row.cells
        for _, cell in self._physical_row_cells(row, cells):
            text = cell.text
            for marker in PLACEHOLDERS:
                if marker in text:
                    self._remove_text(cell, marker)
                    text = cell.text

    def _clear_unused_copies_and_find_markers(
        self,
        copies: list[TemplateCopy],
        used_copy_count: int,
    ) -> list[str]:
        """Limpia y valida cada celda física de la plantilla en un solo pase."""
        if self.document is None:
            return []

        unused_ranges: dict[object, list[tuple[int, int]]] = {}
        for copy in copies[used_copy_count:]:
            unused_ranges.setdefault(copy.table._tbl, []).append(
                (copy.start_row, copy.end_row)
            )

        congregation_marker = "[NOMBRE DE LA CONGREGACIÓN]"
        markers_to_clear = tuple(
            marker for marker in TEMPLATE_MARKERS if marker != congregation_marker
        )
        found: set[str] = set()
        for table in self.document.tables:
            # Solo se cambia texto: las referencias sirven para ambos recorridos.
            rows_cells = [row.cells for row in table.rows]
            cells_by_identity: dict[object, set[int]] = {}
            for row_index, cells in enumerate(rows_cells):
                for cell in cells:
                    cells_by_identity.setdefault(cell._tc, set()).add(row_index)

            ranges = unused_ranges.get(table._tbl, [])
            seen_cells: set[object] = set()
            for cells in rows_cells:
                time_cell = cells[0]._tc if cells else None
                for cell in cells:
                    if cell._tc in seen_cells:
                        continue
                    seen_cells.add(cell._tc)
                    cell_rows = cells_by_identity[cell._tc]
                    is_unused = bool(cell_rows) and all(
                        any(start <= cell_row < end for start, end in ranges)
                        for cell_row in cell_rows
                    )
                    text = cell.text
                    if congregation_marker in text:
                        self._replace_in_cell(
                            cell,
                            congregation_marker,
                            CONGREGATION_NAME,
                        )
                        text = cell.text
                    if is_unused:
                        for marker in markers_to_clear:
                            if marker in text:
                                self._remove_text(cell, marker)
                                text = cell.text
                        if cell._tc == time_cell and text.strip() == "0:00":
                            self._remove_text(cell, "0:00")
                            text = cell.text
                    found.update(
                        marker for marker in TEMPLATE_MARKERS if marker in text
                    )
                    if cell._tc == time_cell and text.strip() == "0:00":
                        found.add("0:00")
        return sorted(found)

    def _remove_unused_structure(
        self,
        copies: list[TemplateCopy],
        weeks: list[MeetingWeek],
    ) -> None:
        """Elimina formularios y filas de asignación que no se usan."""
        copies_by_table: dict[object, list[TemplateCopy]] = {}
        for copy in copies:
            copies_by_table.setdefault(copy.table._tbl, []).append(copy)
        unused_by_table: dict[object, list[TemplateCopy]] = {}
        for copy in copies[len(weeks) :]:
            unused_by_table.setdefault(copy.table._tbl, []).append(copy)

        removed_tables: set[object] = set()
        for unused_copies in unused_by_table.values():
            table = unused_copies[0].table
            if len(unused_copies) == len(copies_by_table[table._tbl]):
                self._remove_table_and_empty_separator(table)
                removed_tables.add(table._tbl)
                continue
            rows = list(table.rows)
            for copy in sorted(
                unused_copies, key=lambda item: item.start_row, reverse=True
            ):
                for row_index in range(copy.end_row - 1, copy.start_row - 1, -1):
                    table._tbl.remove(rows[row_index]._tr)
            if not table.rows:
                self._remove_table_and_empty_separator(table)
                removed_tables.add(table._tbl)

        rows_to_remove: dict[object, tuple[Table, set[int]]] = {}
        for week, copy in zip(weeks, copies):
            if copy.table._tbl in removed_tables:
                continue
            christian_life = self._normal_christian_life_assignments(week)
            unused_rows = (
                copy.structure["ministry_rows"][len(week.ministry) :]
                + copy.structure["cv_normal_rows"][len(christian_life) :]
            )
            table_rows = rows_to_remove.setdefault(copy.table._tbl, (copy.table, set()))
            table_rows[1].update(row.row_index for row in unused_rows)

        for table, row_indexes in rows_to_remove.values():
            rows = list(table.rows)
            for row_index in sorted(row_indexes, reverse=True):
                table._tbl.remove(rows[row_index]._tr)

    def _remove_table_and_empty_separator(self, table: Table) -> None:
        next_element = table._tbl.getnext()
        parent = table._tbl.getparent()
        parent.remove(table._tbl)
        if next_element is not None and not "".join(next_element.itertext()).strip():
            parent.remove(next_element)

    # ==================================================================
    # Reemplazo de texto preservando formato (nivel de runs)
    # ==================================================================

    def _set_cell_text(
        self,
        table: Table,
        row_index: int,
        column_index: int,
        new_text: str,
        old_text: str = "",
    ) -> None:
        rows = list(table.rows)
        if row_index >= len(rows):
            return
        cells = rows[row_index].cells
        if column_index >= len(cells):
            return
        cell = cells[column_index]
        search = old_text if old_text else cell.text.strip()
        if search and search in cell.text:
            self._replace_in_cell(cell, search, new_text)

    def _replace_in_cell_conditional(self, cell: _Cell, old: str, new: str) -> bool:
        """Reemplaza solo si old está presente y new no está vacío."""
        if not new or old not in cell.text:
            return False
        return self._replace_in_cell(cell, old, new)

    def _paragraph_text(self, paragraph: Paragraph) -> str:
        return "".join(run.text for run in paragraph.runs)

    def _leading_number_prefix(self, text: str) -> str:
        stripped = text.lstrip()
        digits: list[str] = []
        for character in stripped:
            if character.isdigit():
                digits.append(character)
                continue
            break
        if not digits:
            return ""
        remaining = stripped[len(digits) :].lstrip()
        if not remaining.startswith("."):
            return ""
        return f"{''.join(digits)}."

    def _first_duration_marker_index(self, text: str) -> int:
        markers = ("[X mins.]", "[XX mins.]", "(X mins.)", "(XX mins.)")
        positions = [text.find(marker) for marker in markers if text.find(marker) != -1]
        if not positions:
            return -1
        return min(positions)

    def _replace_in_cell(self, cell: _Cell, old: str, new: str) -> bool:
        if not new:
            return False
        for paragraph in cell.paragraphs:
            if self._replace_in_paragraph(paragraph, old, new):
                return True
        return False

    def _remove_text(self, cell: _Cell, text: str) -> None:
        for paragraph in cell.paragraphs:
            self._replace_in_paragraph(paragraph, text, "")

    def _replace_in_paragraph(
        self, paragraph: Paragraph, search: str, replacement: str
    ) -> bool:
        runs = paragraph.runs
        if not runs:
            return False
        paragraph_text, run_spans = "", []
        for run_index, run in enumerate(runs):
            if run.text:
                run_start = len(paragraph_text)
                paragraph_text += run.text
                run_spans.append((run_index, run_start, len(paragraph_text)))
        match_start = paragraph_text.find(search)
        if match_start == -1:
            return False
        match_end = match_start + len(search)
        affected_run_indices = []
        for run_index, run_start, run_end in run_spans:
            if run_start <= match_start < run_end:
                affected_run_indices.append(run_index)
            elif match_start <= run_start < match_end:
                affected_run_indices.append(run_index)
            elif (
                run_start < match_end <= run_end
                and run_index not in affected_run_indices
            ):
                affected_run_indices.append(run_index)
        if not affected_run_indices:
            return False
        affected_run_indices.sort()
        first_run_index = affected_run_indices[0]
        first_run_start = next(
            run_start
            for run_index, run_start, run_end in run_spans
            if run_index == first_run_index
        )
        relative_start = match_start - first_run_start
        original_text = runs[first_run_index].text
        runs[first_run_index].text = (
            original_text[:relative_start]
            + replacement
            + original_text[relative_start + len(search) :]
        )
        for run_index in affected_run_indices[1:]:
            run_start = next(
                run_start
                for index, run_start, run_end in run_spans
                if index == run_index
            )
            run_end = next(
                run_end for index, run_start, run_end in run_spans if index == run_index
            )
            if run_start >= match_start and run_end <= match_end:
                runs[run_index].text = ""
            elif run_start < match_end <= run_end:
                overlap_length = match_end - run_start
                if 0 < overlap_length <= len(runs[run_index].text):
                    runs[run_index].text = runs[run_index].text[overlap_length:]
            elif run_start <= match_start < run_end:
                overlap_length = run_end - match_start
                if 0 < overlap_length <= len(runs[run_index].text):
                    runs[run_index].text = runs[run_index].text[:-overlap_length]
        return True


def fill_template(
    template_path: Path,
    weeks: list[MeetingWeek],
    output_path: Path,
) -> GenerationResult:
    writer = TemplateWriter(template_path)
    return writer.fill(weeks, output_path)
