"""
Módulo parser para el Meeting Generator.

Analiza el documento fuente (.docx) del programa de reunión
y extrae la información estructurada de cada semana.

Estructura real del documento:
- Encabezados de semana (SEMANA DEL..., PRESIDENTE:...) en párrafos.
- Contenido de cada semana (puntos) en TABLAS de 3 columnas.
- Las secciones se detectan por sus encabezados, no por rango de números.

El resultado solo se devuelve si todas las semanas cumplen el contrato.
"""

import logging
from pathlib import Path
from typing import Optional

from docx import Document
from docx.document import Document as DocxDocument
from docx.table import Table, _Cell

from .config import (
    CHRISTIAN_LIFE_PRE_STUDY_MINUTES,
    DEFAULT_DURATIONS,
    RE_BIBLE_STUDY,
    RE_CONCLUSION_WORDS,
    RE_DATE,
    RE_FINAL_PRAYER,
    RE_PRESIDENT_SONG,
    RE_SONG,
    RE_WEEK_HEADER,
    RE_WEEKLY_READING,
)
from .models import Assignment, MeetingWeek, NameCorrection, Participant
from .utils import (
    clean_text,
    extract_number,
    normalize_person_name,
    normalize_sentence_case,
    parse_single_name,
    remove_name_suffix,
    split_names,
)
from .validation import GenerationValidationError

logger = logging.getLogger(__name__)


class ProgramParser:
    """Parser del documento de programa de reunión."""

    def __init__(
        self,
        file_path: Path,
        *,
        name_corrections: tuple[NameCorrection, ...] = (),
    ) -> None:
        self.file_path: Path = file_path
        self.document: Optional[DocxDocument] = None
        self.weeks: list[MeetingWeek] = []
        self.detected_week_count = 0
        self._name_corrections = {
            normalize_person_name(correction.original_name): correction
            for correction in name_corrections
        }
        self._name_correction_count = len(name_corrections)

    def parse(self) -> list[MeetingWeek]:
        """Analiza y valida el documento sin devolver resultados parciales."""
        logger.info("Iniciando análisis del documento fuente.")
        self.document = Document(str(self.file_path))
        self.weeks = []
        self.detected_week_count = 0
        issues: list[str] = []

        week_headers = self._extract_week_headers_from_paragraphs()
        logger.info("Encabezados encontrados: %d semanas", len(week_headers))

        tables = self.document.tables
        program_tables = [table for table in tables if self._is_program_table(table)]
        logger.info(
            "Tablas encontradas: %d; tablas de programa: %d",
            len(tables),
            len(program_tables),
        )
        self.detected_week_count = max(len(week_headers), len(program_tables))
        if len(self._name_corrections) != self._name_correction_count:
            raise GenerationValidationError(
                "Correcciones de nombres inválidas",
                ["las correcciones repiten un mismo nombre original normalizado"],
            )
        if not program_tables:
            issues.append("el documento no contiene tablas de semanas")
        if len(week_headers) != len(program_tables):
            issues.append(
                "la cantidad de encabezados "
                f"({len(week_headers)}) no coincide con la de tablas de programa "
                f"({len(program_tables)})"
            )

        for week_index, table in enumerate(program_tables):
            try:
                header_data = (
                    week_headers[week_index] if week_index < len(week_headers) else {}
                )
                meeting_week = self._parse_week_from_table(
                    table,
                    week_index + 1,
                    header_data,
                )

                logger.info(
                    "Semana %d: tesoros=%d ministerio=%d vida_cristiana=%d",
                    week_index + 1,
                    len(meeting_week.bible_treasures),
                    len(meeting_week.ministry),
                    len(meeting_week.christian_life),
                )
                self.weeks.append(meeting_week)

            except GenerationValidationError as exc:
                logger.error("Error tabla %d: %s", week_index + 1, exc)
                issues.extend(exc.issues)
            except Exception as exc:
                logger.error("Error tabla %d: %s", week_index + 1, exc, exc_info=True)
                issues.append(
                    f"tabla {week_index + 1}: no se pudo analizar "
                    f"({type(exc).__name__}: {exc})"
                )

        for week in self.weeks:
            if not week.is_valid():
                issues.extend(
                    f"semana {week.week_index}: {issue}"
                    for issue in week.validation_errors()
                )

        if issues:
            raise GenerationValidationError("Documento fuente inválido", issues)

        logger.info("Análisis completado. %d semanas extraídas.", len(self.weeks))
        return self.weeks

    def _is_program_table(self, table: Table) -> bool:
        """Distingue una tabla de semana de una tabla auxiliar."""
        sections: set[str] = set()
        for row in table.rows:
            cells = row.cells
            if len(cells) > 1:
                section = self._detect_section_from_lines(
                    self._split_cell_lines(cells[1].text)
                )
                if section:
                    sections.add(section)
        return sections == {"treasures", "ministry", "christian_life"}

    # ------------------------------------------------------------------
    # Encabezados desde párrafos
    # ------------------------------------------------------------------

    def _extract_week_headers_from_paragraphs(self) -> list[dict[str, str]]:
        """Extrae encabezados de semana desde los párrafos."""
        if self.document is None:
            return []

        paragraphs = [paragraph.text for paragraph in self.document.paragraphs]
        headers: list[dict[str, str]] = []
        current_header: dict[str, str] = {}

        for paragraph_text in paragraphs:
            stripped = paragraph_text.strip()
            if not stripped:
                continue

            if RE_WEEK_HEADER.match(stripped):
                if current_header:
                    headers.append(current_header)
                current_header = {"raw": stripped}
                date, weekly_reading = self._extract_week_header_fields(stripped)
                current_header["date"] = date
                if weekly_reading:
                    current_header["weekly_reading"] = weekly_reading
                continue

            president_match = RE_PRESIDENT_SONG.search(stripped)
            if president_match and "president" not in current_header:
                current_header["president"] = normalize_person_name(
                    president_match.group(1)
                )
                current_header["opening_song"] = clean_text(president_match.group(2))
                continue

            if "weekly_reading" not in current_header and current_header:
                reading_match = RE_WEEKLY_READING.search(stripped)
                if reading_match:
                    current_header["weekly_reading"] = self._normalize_weekly_reading(
                        reading_match.group(0)
                    )

        if current_header:
            headers.append(current_header)
        return headers

    @staticmethod
    def _extract_week_header_fields(header: str) -> tuple[str, str]:
        """Separa fecha y lectura usando la posición real de la lectura."""
        cleaned_header = clean_text(header)
        reading_match = RE_WEEKLY_READING.search(cleaned_header)
        date_end = reading_match.start() if reading_match else len(cleaned_header)
        date_match = RE_DATE.search(cleaned_header[:date_end])
        date_text = date_match.group(1) if date_match else ""
        date = normalize_sentence_case(date_text.rstrip(" \t.,;:-–—"))
        weekly_reading = (
            ProgramParser._normalize_weekly_reading(reading_match.group(0))
            if reading_match
            else ""
        )
        return date, weekly_reading

    @staticmethod
    def _normalize_weekly_reading(reading: str) -> str:
        """Normaliza la lectura y conserva la inicial tras un ordinal."""
        normalized = normalize_sentence_case(reading)
        for index, character in enumerate(normalized):
            if character.isalpha():
                return normalized[:index] + character.upper() + normalized[index + 1 :]
        return normalized

    # ------------------------------------------------------------------
    # Parseo de semana desde tabla
    # ------------------------------------------------------------------

    def _parse_week_from_table(
        self, table: Table, week_number: int, header_data: dict[str, str]
    ) -> MeetingWeek:
        """
        Parsea una semana desde una tabla del documento.

        Estrategia dinámica:
        - Mantiene la sección actual ("treasures", "ministry", "christian_life").
        - Cambia de sección al encontrar encabezados.
        - Cada punto se convierte en un Assignment con su número original.
        """
        week = MeetingWeek(week_index=week_number)

        week.raw_header = header_data.get("raw", "")
        week.date = header_data.get("date", "")
        week.weekly_reading = header_data.get("weekly_reading", "")
        week.president = header_data.get("president", "")
        if week.president:
            week.president = self._validate_participant_text(
                week.president, f"semana {week_number}, presidente", single=True
            )
        week.opening_song = header_data.get("opening_song", "")
        week.opening_prayer = week.president

        current_section: str = "treasures"

        for row_number, row in enumerate(table.rows, start=1):
            cells = row.cells
            if not any(clean_text(cell.text) for cell in cells):
                continue
            context = f"semana {week_number}, fila {row_number}"
            omitted_columns = row._tr.xpath(
                "w:trPr/w:gridBefore/@w:val | w:trPr/w:gridAfter/@w:val"
            )
            if len(cells) != 3 or any(int(value) > 0 for value in omitted_columns):
                raise GenerationValidationError(
                    "Fila de programa incompleta",
                    [
                        f"{context}: se requieren las columnas de número, contenido "
                        "y participantes. No se puede interpretar con seguridad."
                    ],
                )
            if len({cell._tc for cell in cells}) != 3:
                merged_lines = self._split_cell_lines(
                    "\n".join(cell.text for cell in cells)
                )
                if all(
                    self._detect_section_from_lines([line]) or RE_SONG.match(line)
                    for line in merged_lines
                ):
                    new_section = self._detect_section_from_lines(merged_lines)
                    if new_section:
                        current_section = new_section
                    self._process_special_row(
                        week, self._split_cell_lines(cells[1].text), ""
                    )
                    continue
                raise GenerationValidationError(
                    "Fila de programa ambigua",
                    [f"{context}: las columnas de una asignación no pueden fusionarse"],
                )

            point_numbers = self._extract_point_numbers(cells[0])
            point_number = point_numbers[0] if len(point_numbers) == 1 else None
            content_text = cells[1].text
            lines = self._split_cell_lines(content_text)
            person_text = clean_text(cells[2].text)

            new_section = self._detect_section_from_lines(lines)
            if new_section:
                current_section = new_section

            # Algunas tablas agrupan varias asignaciones en una misma fila.
            # En ese caso los números, títulos y participantes se conservan en
            # párrafos separados y deben procesarse individualmente.
            participant_groups = sum(
                bool(clean_text(paragraph.text)) for paragraph in cells[2].paragraphs
            )
            if len(point_numbers) > 1 or participant_groups > 1:
                assignments = self._create_assignments_from_combined_row(
                    point_numbers,
                    cells[1],
                    cells[2],
                    week_number=week_number,
                    row_number=row_number,
                )
                self._append_assignments(week, current_section, assignments)
                continue

            is_final_prayer = self._lines_contain_final_prayer(lines)
            primary_name = self._validate_participant_text(
                person_text,
                context,
                single=is_final_prayer and bool(person_text),
            )
            assignment = self._create_assignment(
                point_number or 0, lines, person_text, context=context
            )
            has_assignment_content = bool(
                assignment.title.strip()
                or assignment.duration_text
                or assignment.has_participants()
            )
            if (
                cells[0]._tc.getparent() is not row._tr
                and has_assignment_content
                and not is_final_prayer
            ):
                raise GenerationValidationError(
                    "Fila de programa ambigua",
                    [f"{context}: el número está fusionado con otra fila"],
                )
            if is_final_prayer and (
                assignment.title.strip() or assignment.duration_text
            ):
                raise GenerationValidationError(
                    "Fila de cierre ambigua",
                    [f"{context}: la oración final comparte fila con una asignación"],
                )

            if self._lines_contain_conclusion(lines):
                conclusion_lines = [
                    line for line in lines if RE_CONCLUSION_WORDS.match(line)
                ]
                conclusion = self._create_assignment(
                    0, conclusion_lines, "", context=context
                )
                remaining = self._create_assignment(
                    0,
                    [line for line in lines if not RE_CONCLUSION_WORDS.match(line)],
                    "",
                    context=context,
                )
                if (
                    len(conclusion_lines) != 1
                    or conclusion.title.upper().rstrip(".,;: ")
                    not in ("PALABRAS DE CONCLUSIÓN", "PALABRAS DE CONCLUSION")
                    or remaining.title.strip()
                    or (remaining.duration_text and conclusion.duration_text)
                ):
                    raise GenerationValidationError(
                        "Fila de cierre ambigua",
                        [f"{context}: la conclusión comparte fila con otra asignación"],
                    )
                assignment.is_conclusion = True
                week.christian_life.append(assignment)
                continue

            if is_final_prayer:
                if person_text:
                    week.closing_prayer = normalize_person_name(
                        remove_name_suffix(primary_name)
                    )
                continue

            if point_number is None:
                if (
                    current_section in ("treasures", "ministry")
                    and has_assignment_content
                ):
                    section_name = (
                        "Tesoros de la Biblia"
                        if current_section == "treasures"
                        else "Seamos Mejores Maestros"
                    )
                    raise GenerationValidationError(
                        "Asignación sin número",
                        [
                            f"{context}: falta el "
                            f"número de una asignación de {section_name}. "
                            "No se puede inferir con seguridad."
                        ],
                    )
                if has_assignment_content:
                    if not assignment.title.strip():
                        raise GenerationValidationError(
                            "Asignación incompleta",
                            [f"{context}: falta el título de la asignación"],
                        )
                    self._append_assignments(
                        week,
                        current_section,
                        [assignment],
                    )
                    continue
                self._process_special_row(week, lines, person_text)
                continue

            # Hay documentos donde el encabezado de Vida Cristiana y su canción
            # aparecen en una fila que conserva por error un número de punto.
            if not assignment.title.strip():
                if (
                    assignment.duration_text
                    or assignment.has_participants()
                    or not lines
                ):
                    raise GenerationValidationError(
                        "Asignación incompleta",
                        [f"{context}: falta el título de la asignación"],
                    )
                self._process_special_row(week, lines, person_text)
                continue

            self._append_assignments(week, current_section, [assignment])

        self._extract_closing_from_table(week, table)
        self._assign_missing_christian_life_numbers(week)
        self._infer_missing_ministry_durations(week)
        self._infer_missing_christian_life_duration(week)
        self._apply_name_corrections(week)

        return week

    def _extract_point_numbers(self, cell: _Cell) -> list[int]:
        """Extrae por separado los números escritos en párrafos de una celda."""
        numbers: list[int] = []
        for paragraph in cell.paragraphs:
            number = self._clean_point_number(clean_text(paragraph.text))
            if number is not None:
                numbers.append(number)
        return numbers

    def _create_assignments_from_combined_row(
        self,
        point_numbers: list[int],
        content_cell: _Cell,
        participant_cell: _Cell,
        *,
        week_number: int,
        row_number: int,
    ) -> list[Assignment]:
        """Separa una fila que contiene más de una asignación."""
        content_lines = [
            clean_text(paragraph.text)
            for paragraph in content_cell.paragraphs
            if clean_text(paragraph.text)
        ]
        participant_lines = [
            clean_text(paragraph.text)
            for paragraph in participant_cell.paragraphs
            if clean_text(paragraph.text)
        ]

        if not len(point_numbers) == len(content_lines) == len(participant_lines):
            raise GenerationValidationError(
                "Fila combinada ambigua",
                [
                    f"semana {week_number}, fila {row_number}: fila combinada "
                    f"ambigua (números={len(point_numbers)}, "
                    f"títulos={len(content_lines)}, "
                    f"grupos de participantes={len(participant_lines)}). "
                    "Cada número debe tener un título y un grupo de participantes."
                ],
            )

        for participant_text in participant_lines:
            self._validate_participant_text(
                participant_text, f"semana {week_number}, fila {row_number}"
            )

        return [
            self._create_assignment(
                number,
                [title_line],
                participant_text,
                context=f"semana {week_number}, fila {row_number}",
            )
            for number, title_line, participant_text in zip(
                point_numbers, content_lines, participant_lines
            )
        ]

    def _append_assignments(
        self,
        week: MeetingWeek,
        section: str,
        assignments: list[Assignment],
    ) -> None:
        """Agrega asignaciones a la sección activa."""
        if section == "treasures":
            week.bible_treasures.extend(assignments)
        elif section == "ministry":
            week.ministry.extend(assignments)
        elif section == "christian_life":
            week.christian_life.extend(assignments)

    def _infer_missing_christian_life_duration(self, week: MeetingWeek) -> None:
        """Completa una única duración omitida antes del estudio bíblico."""
        assignments = [
            assignment
            for assignment in week.christian_life
            if not assignment.is_bible_study
            and not assignment.is_conclusion
            and assignment.title.strip()
        ]
        missing = [
            assignment for assignment in assignments if assignment.duration_minutes <= 0
        ]
        known_minutes = sum(
            assignment.duration_minutes
            for assignment in assignments
            if assignment.duration_minutes > 0
        )
        remaining = CHRISTIAN_LIFE_PRE_STUDY_MINUTES - known_minutes

        if len(missing) == 1 and remaining > 0:
            missing[0].duration_minutes = remaining
            missing[0].duration_text = f"{remaining} mins."
            logger.info(
                "Semana %d, punto %d: duración inferida de %d minutos.",
                week.week_index,
                missing[0].number,
                remaining,
            )
        elif missing:
            logger.warning(
                "Semana %d: no se pudo inferir con seguridad la duración de %s.",
                week.week_index,
                [assignment.number for assignment in missing],
            )

    def _assign_missing_christian_life_numbers(self, week: MeetingWeek) -> None:
        """Asigna números omitidos usando la posición anterior al estudio."""
        unnumbered = [
            assignment
            for assignment in week.christian_life
            if assignment.number <= 0
            and not assignment.is_bible_study
            and not assignment.is_conclusion
        ]
        bible_study = next(
            (
                assignment
                for assignment in week.christian_life
                if assignment.is_bible_study and assignment.number > 0
            ),
            None,
        )
        if not unnumbered or bible_study is None:
            return

        used_numbers = {
            assignment.number
            for assignment in week.all_assignments()
            if assignment.number > 0
        }
        available_numbers = [
            number
            for number in range(1, bible_study.number)
            if number not in used_numbers
        ]
        inferred_numbers = available_numbers[-len(unnumbered) :]
        if len(inferred_numbers) != len(unnumbered):
            logger.warning(
                "Semana %d: no se pudieron inferir números para %d asignaciones.",
                week.week_index,
                len(unnumbered),
            )
            return

        for assignment, number in zip(unnumbered, inferred_numbers):
            assignment.number = number
            logger.info(
                "Semana %d: número %d inferido para una asignación de Vida Cristiana.",
                week.week_index,
                number,
            )

    def _infer_missing_ministry_durations(self, week: MeetingWeek) -> None:
        """Aplica la duración estándar a discursos sin tiempo explícito."""
        for assignment in week.ministry:
            if assignment.duration_minutes <= 0 and assignment.title.upper().startswith(
                "DISCURSO"
            ):
                duration = DEFAULT_DURATIONS["ministerio_discurso"]
                assignment.duration_minutes = duration
                assignment.duration_text = f"{duration} mins."
                logger.info(
                    "Semana %d, punto %d: duración de discurso inferida en %d minutos.",
                    week.week_index,
                    assignment.number,
                    duration,
                )

    # ------------------------------------------------------------------
    # Detección de secciones por encabezado
    # ------------------------------------------------------------------

    def _detect_section_from_lines(self, lines: list[str]) -> Optional[str]:
        """
        Detecta a qué sección pertenece una fila según sus líneas de texto.
        Retorna "treasures", "ministry", "christian_life", o None.
        """
        for line in lines:
            upper = line.upper().strip().rstrip(".")
            if upper in ("TESOROS DE LA BIBLIA", "TESOROS DE LAS BIBLIAS"):
                return "treasures"
            if upper == "SEAMOS MEJORES MAESTROS":
                return "ministry"
            if upper == "NUESTRA VIDA CRISTIANA":
                return "christian_life"
        return None

    def _lines_contain_conclusion(self, lines: list[str]) -> bool:
        """Verifica si las líneas contienen 'Palabras de conclusión'."""
        for line in lines:
            if RE_CONCLUSION_WORDS.match(line):
                return True
        return False

    def _lines_contain_final_prayer(self, lines: list[str]) -> bool:
        """Verifica si las líneas contienen 'Oración final'."""
        for line in lines:
            if RE_FINAL_PRAYER.match(line):
                return True
        return False

    # ------------------------------------------------------------------
    # Creación de Assignment unificada
    # ------------------------------------------------------------------

    def _create_assignment(
        self,
        point_number: int,
        lines: list[str],
        person_text: str,
        *,
        context: str = "",
    ) -> Assignment:
        """Extrae el título, la duración y los participantes de una asignación."""
        import re as re_module

        assignment = Assignment(number=point_number)
        assignment.raw_lines = lines

        is_bible_study = False
        is_conclusion = False
        is_prayer = False

        for line in lines:
            if RE_BIBLE_STUDY.match(line):
                is_bible_study = True
            if RE_CONCLUSION_WORDS.match(line):
                is_conclusion = True
            if RE_FINAL_PRAYER.match(line):
                is_prayer = True

        assignment.is_bible_study = is_bible_study
        assignment.is_conclusion = is_conclusion
        assignment.is_prayer = is_prayer

        content_lines = []
        for line in lines:
            if self._detect_section_from_lines([line]):
                continue
            if RE_SONG.match(line):
                continue
            if RE_FINAL_PRAYER.match(line):
                continue
            cleaned = clean_text(line)
            if cleaned:
                content_lines.append(cleaned)

        full_text = " ".join(content_lines).strip()
        full_text = re_module.sub(r"\s+", " ", full_text)

        if not full_text:
            logger.debug(
                "Assignment %d: sin contenido textual (solo encabezado).", point_number
            )
            if person_text:
                assignment.participants = self._parse_participants(
                    person_text, is_bible_study
                )
            return assignment

        # La duración puede ir seguida de aclaraciones como "(parte 1)".
        duration_pattern = re_module.compile(
            r"[\(\.]?\s*(\d+)\s*(?:minutos?|mins?\.?)\s*\)?\.?",
            re_module.IGNORECASE,
        )
        duration_matches = list(duration_pattern.finditer(full_text))
        if len(duration_matches) > 1:
            prefix = f"{context}: " if context else ""
            raise GenerationValidationError(
                "Fila combinada ambigua",
                [
                    f"{prefix}hay varias duraciones de asignación sin sus "
                    "correspondientes números y grupos de participantes"
                ],
            )
        duration_match = duration_matches[-1] if duration_matches else None

        if duration_match:
            duration_minutes = int(duration_match.group(1))
            assignment.duration_minutes = duration_minutes
            assignment.duration_text = f"{duration_minutes} mins."

            # Las aclaraciones posteriores forman parte del título de la fuente.
            title = clean_text(
                full_text[: duration_match.start()] + full_text[duration_match.end() :]
            ).strip("., ")
            assignment.title = self._shorten_activity_title(title, duration_minutes)
        else:
            assignment.title = self._shorten_activity_title(full_text)

        if not assignment.title.strip():
            logger.warning(
                "Punto %d: título vacío (estudio_bíblico=%s, conclusión=%s).",
                point_number,
                is_bible_study,
                is_conclusion,
            )
        if not assignment.duration_text:
            logger.debug(
                "Punto %d: sin duración detectada.",
                point_number,
            )

        if person_text:
            assignment.participants = self._parse_participants(
                person_text, is_bible_study
            )

        return assignment

    def _parse_participants(
        self, person_text: str, is_bible_study: bool = False
    ) -> list[Participant]:
        """
        Parsea los participantes desde el texto de la columna 2.

        Divide por '//' y asigna roles automáticamente.
        """
        participants: list[Participant] = []

        if not person_text:
            return participants

        primary_name, secondary_name = split_names(person_text)

        if is_bible_study:
            if primary_name:
                participants.append(Participant(name=primary_name, role="Conductor"))
            if secondary_name:
                participants.append(Participant(name=secondary_name, role="Lector"))
        else:
            if primary_name:
                participants.append(Participant(name=primary_name, role="Estudiante"))
            if secondary_name:
                participants.append(Participant(name=secondary_name, role="Ayudante"))

        return participants

    @staticmethod
    def _validate_participant_text(
        text: str, context: str, *, single: bool = False
    ) -> str:
        """Valida el grupo completo y conserva la ubicación de cualquier error."""
        try:
            return parse_single_name(text) if single else split_names(text)[0]
        except ValueError as exc:
            raise GenerationValidationError(
                "Participantes inválidos", [f"{context}: {exc}"]
            ) from exc

    @staticmethod
    def _shorten_activity_title(title: str, duration_minutes: int = 0) -> str:
        """Conserva el título y descarta instrucciones de la asignación."""
        import re as re_module

        detail = re_module.search(
            r"(?:\s*,\s*|\s*\.\s*|\s+)"
            r"(?:DE\s+CASA\s+EN\s+CASA|PREDICACI[ÓO]N\s+"
            r"(?:IN\s*FORMAL|INFORMAL|PÚBLICA)|ESCENIFICACIÓN|"
            r"CONVERSE\s+CON|LA\s+PERSONA|"
            r"EMPIECE\s+UNA\s+CONVERSACIÓN|JESÚS\s+NO\s+ES\s+DIOS)\b",
            title,
            re_module.IGNORECASE,
        )
        if title.startswith("“"):
            detail = re_module.search(
                r"(?<=[”»])\.\s*ANÁLISIS\s+CON\b",
                title,
                re_module.IGNORECASE,
            )
        elif detail is None and duration_minutes in {6, 9}:
            detail = re_module.search(
                r"\.\s*ANÁLISIS\s+CON\b", title, re_module.IGNORECASE
            )
        shortened = title[: detail.start()] if detail else title
        shortened = shortened.replace("(….)", "[…]")
        shortened = re_module.sub(r"“\s+", "“", shortened)
        shortened = re_module.sub(r"\s+([”»])", r"\1", shortened)
        shortened = re_module.sub(r"\s+“$", "”", shortened)
        return shortened.capitalize() if shortened.isupper() else shortened

    def _apply_name_corrections(self, week: MeetingWeek) -> None:
        week.president = self._correct_person_name(week.president, week, "presidente")
        week.opening_prayer = self._correct_person_name(
            week.opening_prayer, week, "oración inicial"
        )
        week.closing_prayer = self._correct_person_name(
            week.closing_prayer, week, "oración final"
        )
        for assignment in week.all_assignments():
            for index, participant in enumerate(assignment.participants, start=1):
                participant.name = self._correct_person_name(
                    participant.name,
                    week,
                    f"punto {assignment.number}, participante {index}",
                )

    def _correct_person_name(self, name: str, week: MeetingWeek, field: str) -> str:
        correction = self._name_corrections.get(name)
        if correction is None:
            return name
        corrected_name = normalize_person_name(correction.corrected_name)
        if corrected_name != name:
            logger.info(
                "Corrección explícita de nombre: fuente=%s, semana=%d, campo=%s, "
                "original=%r, corregido=%r, motivo=%s",
                self.file_path,
                week.week_index,
                field,
                name,
                corrected_name,
                clean_text(correction.reason),
            )
        return corrected_name

    # ------------------------------------------------------------------
    # Utilidades de la tabla
    # ------------------------------------------------------------------

    def _clean_point_number(self, text: str) -> Optional[int]:
        """Limpia el texto de número de punto."""
        if not text:
            return None
        cleaned = text.replace("//", "").strip()
        return extract_number(cleaned)

    def _split_cell_lines(self, cell_text: str) -> list[str]:
        """Divide el texto de una celda en líneas significativas."""
        lines = cell_text.split("\n")
        result = []
        for line in lines:
            cleaned = clean_text(line)
            if cleaned:
                result.append(cleaned)
        return result

    def _process_special_row(
        self, week: MeetingWeek, lines: list[str], person_text: str
    ) -> None:
        """Procesa una fila sin número de punto (solo encabezados/canciones)."""
        for line in lines:
            song_match = RE_SONG.match(line)
            if song_match:
                song_num = song_match.group(1)
                if not week.intermediate_song:
                    week.intermediate_song = song_num
                elif not week.closing_song:
                    week.closing_song = song_num

    # ------------------------------------------------------------------
    # Extracción del cierre
    # ------------------------------------------------------------------

    def _extract_closing_from_table(self, week: MeetingWeek, table: Table) -> None:
        """Extrae canción final y oración final del final de la tabla."""
        songs: list[str] = []
        for row_number, row in enumerate(table.rows, start=1):
            cells = row.cells
            if len(cells) < 2:
                continue

            content_text = cells[1].text
            person_text = clean_text(cells[2].text) if len(cells) > 2 else ""

            lines = self._split_cell_lines(content_text)
            for line in lines:
                if RE_FINAL_PRAYER.match(line):
                    if person_text:
                        student = self._validate_participant_text(
                            person_text,
                            f"semana {week.week_index}, fila {row_number}",
                            single=True,
                        )
                        if not week.closing_prayer:
                            week.closing_prayer = normalize_person_name(
                                remove_name_suffix(student)
                            )
                    continue

                song_match = RE_SONG.match(line)
                if song_match:
                    songs.append(song_match.group(1))

        if songs:
            week.intermediate_song = songs[0]
        if len(songs) > 1:
            week.closing_song = songs[-1]


def parse_document(
    file_path: Path,
    *,
    name_corrections: tuple[NameCorrection, ...] = (),
) -> list[MeetingWeek]:
    """Función de conveniencia para parsear un documento."""
    parser = ProgramParser(file_path, name_corrections=name_corrections)
    return parser.parse()
