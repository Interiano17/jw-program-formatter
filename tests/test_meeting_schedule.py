import copy
import os
import tempfile
import tkinter as tk
import unittest
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor, wait
from pathlib import Path
from threading import Event, get_ident
from unittest.mock import Mock, patch

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from meeting_generator.config import BIBLE_BOOK_NAMES
from meeting_generator.generator import generate_document
from meeting_generator.settings import AppSettings
from meeting_generator.main import (
    COLOR_ERROR,
    COLOR_MUTED,
    COLOR_SUCCESS,
    MeetingGeneratorApp,
)
from meeting_generator.models import Assignment, NameCorrection
from meeting_generator.parser import ProgramParser, parse_document
from meeting_generator.results import GenerationResult
from meeting_generator.template_writer import TemplateWriter, fill_template
from meeting_generator.utils import (
    normalize_person_name,
    normalize_sentence_case,
    split_names,
)
from meeting_generator.validation import GenerationValidationError
from tests.source_fixtures import (
    build_primary_source,
    build_secondary_source,
    build_source_with_auxiliary_table,
)
from tests.template_fixture import build_template

_FIXTURE_DIRECTORY: tempfile.TemporaryDirectory[str] | None = None
_CONGREGATION_PATCHER = None
SOURCE_DOCUMENT = Path("primary-source-not-built.docx")
SECOND_SOURCE_DOCUMENT = Path("secondary-source-not-built.docx")
TEMPLATE_DOCUMENT = Path("template-not-built.docx")


def _first_form_rows(table):
    starts = [
        index
        for index, row in enumerate(table.rows)
        if any(cell.text.strip() == "Presidente:" for cell in row.cells)
    ]
    return table.rows[starts[0] : starts[1] if len(starts) > 1 else len(table.rows)]


def setUpModule() -> None:
    """Construye todos los DOCX de prueba sin depender del árbol local."""
    global _CONGREGATION_PATCHER
    global _FIXTURE_DIRECTORY
    global SECOND_SOURCE_DOCUMENT
    global SOURCE_DOCUMENT
    global TEMPLATE_DOCUMENT

    _FIXTURE_DIRECTORY = tempfile.TemporaryDirectory()
    fixture_directory = Path(_FIXTURE_DIRECTORY.name)
    SOURCE_DOCUMENT = build_primary_source(fixture_directory / "primary-source.docx")
    SECOND_SOURCE_DOCUMENT = build_secondary_source(
        fixture_directory / "secondary-source.docx"
    )
    TEMPLATE_DOCUMENT = build_template(fixture_directory / "template.docx")
    _CONGREGATION_PATCHER = patch(
        "meeting_generator.template_writer.CONGREGATION_NAME",
        "Congregación de Prueba",
    )
    _CONGREGATION_PATCHER.start()


def tearDownModule() -> None:
    """Elimina los DOCX sintéticos creados para este módulo."""
    if _CONGREGATION_PATCHER is not None:
        _CONGREGATION_PATCHER.stop()
    if _FIXTURE_DIRECTORY is not None:
        _FIXTURE_DIRECTORY.cleanup()


class ParserRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.weeks = parse_document(SOURCE_DOCUMENT)

    def test_numbered_section_song_is_not_an_assignment(self) -> None:
        week = self.weeks[5]

        normal_assignments = [
            assignment
            for assignment in week.christian_life
            if not assignment.is_bible_study and not assignment.is_conclusion
        ]

        self.assertEqual([assignment.number for assignment in normal_assignments], [8])
        self.assertEqual(normal_assignments[0].duration_minutes, 15)

    def test_combined_row_is_split_and_missing_duration_is_inferred(self) -> None:
        week = self.weeks[7]
        normal_assignments = [
            assignment
            for assignment in week.christian_life
            if not assignment.is_bible_study and not assignment.is_conclusion
        ]

        self.assertEqual(
            [assignment.number for assignment in normal_assignments], [7, 8]
        )
        self.assertEqual(
            [assignment.duration_minutes for assignment in normal_assignments],
            [10, 5],
        )
        self.assertEqual(normal_assignments[1].duration_text, "5 mins.")

    def test_combined_row_rejects_mismatched_counts_with_week_and_row(self) -> None:
        cases = (
            ("número adicional", 0, "9", (3, 2, 2)),
            ("título adicional", 1, "Tema adicional", (2, 3, 2)),
            ("título ausente", 1, None, (2, 1, 2)),
            ("participantes adicionales", 2, "Ana Ejemplo", (2, 2, 3)),
            ("participantes ausentes", 2, None, (2, 2, 1)),
            ("encabezado adicional", 1, "NUESTRA VIDA CRISTIANA", (2, 3, 2)),
            ("conclusión adicional", 1, "PALABRAS DE CONCLUSIÓN", (2, 3, 2)),
            ("oración adicional", 1, "ORACIÓN FINAL", (2, 3, 2)),
        )
        with tempfile.TemporaryDirectory() as directory:
            source_path = Path(directory) / "ambiguous-source.docx"
            for label, column, extra_text, counts in cases:
                with self.subTest(case=label):
                    document = Document(SOURCE_DOCUMENT)
                    cell = document.tables[7].rows[10].cells[column]
                    if extra_text is None:
                        cell.paragraphs[-1].text = ""
                    else:
                        cell.add_paragraph(extra_text)
                    document.save(source_path)

                    with self.assertRaises(GenerationValidationError) as raised:
                        parse_document(source_path)

                    expected_counts = (
                        f"números={counts[0]}, títulos={counts[1]}, "
                        f"grupos de participantes={counts[2]}"
                    )
                    self.assertTrue(
                        any(
                            "semana 8, fila 11" in issue and expected_counts in issue
                            for issue in raised.exception.issues
                        ),
                        raised.exception.issues,
                    )

    def test_intermediate_and_closing_songs_are_distinct(self) -> None:
        first_week = self.weeks[0]

        self.assertEqual(first_week.intermediate_song, "49")
        self.assertEqual(first_week.closing_song, "61")

    def test_unnumbered_assignments_are_rejected_with_week_and_row(self) -> None:
        cases = (
            ("treasures", "TESOROS DE LA BIBLIA", "SEAMOS MEJORES MAESTROS"),
            ("ministry", "SEAMOS MEJORES MAESTROS", "NUESTRA VIDA CRISTIANA"),
        )
        with tempfile.TemporaryDirectory() as directory:
            source_path = Path(directory) / "unnumbered-source.docx"
            for section, heading, next_heading in cases:
                for merged_with_heading in (False, True):
                    for content, participants in (
                        ("Asignación adicional (5 mins.)", "Ana Ejemplo"),
                        ("Asignación adicional", "Ana Ejemplo"),
                        ("Asignación adicional (5 mins.)", ""),
                        ("Asignación adicional", ""),
                    ):
                        with self.subTest(
                            section=section,
                            merged_with_heading=merged_with_heading,
                            content=content,
                            participants=participants,
                        ):
                            document = Document(SOURCE_DOCUMENT)
                            table = document.tables[1]
                            if merged_with_heading:
                                row = next(
                                    row
                                    for row in table.rows
                                    if row.cells[1].text == heading
                                )
                                row.cells[1].add_paragraph(content)
                            else:
                                boundary = next(
                                    row
                                    for row in table.rows
                                    if row.cells[1].text == next_heading
                                )
                                row = table.add_row()
                                row.cells[0].text = " \t "
                                row.cells[1].text = content
                                boundary._tr.addprevious(row._tr)
                            row.cells[2].text = participants
                            row_number = next(
                                index
                                for index, candidate in enumerate(table.rows, start=1)
                                if candidate._tr is row._tr
                            )
                            document.save(source_path)

                            with self.assertRaises(GenerationValidationError) as raised:
                                parse_document(source_path)

                            self.assertTrue(
                                any(
                                    f"semana 2, fila {row_number}" in issue
                                    and "falta el número" in issue
                                    for issue in raised.exception.issues
                                ),
                                raised.exception.issues,
                            )

    def test_unnumbered_headings_songs_and_empty_rows_remain_valid(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source_path = Path(directory) / "special-rows.docx"
            document = Document(SOURCE_DOCUMENT)
            table = document.tables[0]
            heading_index, heading = next(
                (index, row)
                for index, row in enumerate(table.rows)
                if row.cells[1].text == "NUESTRA VIDA CRISTIANA"
            )
            song = table.rows[heading_index + 1]
            heading.cells[1].add_paragraph(song.cells[1].text)
            table._tbl.remove(song._tr)
            empty_row = table.add_row()
            heading._tr.addprevious(empty_row._tr)
            document.save(source_path)

            weeks = parse_document(source_path)

        self.assertEqual(weeks, self.weeks)

    def test_auxiliary_table_does_not_shift_week_headers(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source_path = Path(directory) / "auxiliary-table.docx"
            build_source_with_auxiliary_table(source_path)

            weeks = parse_document(source_path)

        self.assertEqual(len(weeks), 2)
        self.assertEqual(weeks[0].weekly_reading, "Isaías 1-3")
        self.assertEqual(weeks[1].weekly_reading, "Oseas 1-3")

    def test_title_preserves_text_after_duration(self) -> None:
        assignment = ProgramParser(Path("unused.docx"))._create_assignment(
            1,
            ["Tema (5 minutos) (parte 1)"],
            "",
        )

        self.assertEqual(assignment.duration_minutes, 5)
        self.assertEqual(assignment.title, "Tema (parte 1)")

    def test_assignment_instructions_are_not_part_of_the_title(self) -> None:
        assignment = ProgramParser(Path("unused.docx"))._create_assignment(
            4,
            [
                "Empiece conversaciones, DE CASA EN CASA. "
                "Ofrezca un curso de la Biblia (3 minutos)"
            ],
            "",
        )

        self.assertEqual(assignment.title, "Empiece conversaciones")


class WeekHeaderParsingTests(unittest.TestCase):
    def _extract_header(self, text: str) -> dict:
        parser = ProgramParser(Path("unused.docx"))
        document = Document()
        document.add_paragraph(text)
        parser.document = document

        headers = parser._extract_week_headers_from_paragraphs()

        self.assertEqual(len(headers), 1)
        return headers[0]

    def test_isaias_reading_is_not_included_in_date(self) -> None:
        header = self._extract_header("SEMANA DEL 5 AL 11 DE ENERO. ISAÍAS 1-3")

        self.assertEqual(header["date"], "5 al 11 de enero")
        self.assertEqual(header["weekly_reading"], "Isaías 1-3")

    def test_numbered_book_preserves_ordinal(self) -> None:
        header = self._extract_header("SEMANA DEL 12 AL 18 DE ENERO. 1 CORINTIOS 1-2")

        self.assertEqual(header["date"], "12 al 18 de enero")
        self.assertEqual(header["weekly_reading"], "1 Corintios 1-2")

    def test_oseas_is_recognized_as_weekly_reading(self) -> None:
        header = self._extract_header("SEMANA DEL 19 AL 25 DE ENERO. OSEAS 1-3")

        self.assertEqual(header["date"], "19 al 25 de enero")
        self.assertEqual(header["weekly_reading"], "Oseas 1-3")

    def test_representative_numbered_and_compound_books_are_recognized(
        self,
    ) -> None:
        readings = (
            ("2 CRÓNICAS 1-3", "2 Crónicas 1-3"),
            (
                "CANTAR DE LOS CANTARES 1-2",
                "Cantar de los cantares 1-2",
            ),
            (
                "EL CANTAR DE LOS CANTARES 1-2",
                "El cantar de los cantares 1-2",
            ),
            (
                "HECHOS DE LOS APÓSTOLES 1-2",
                "Hechos de los apóstoles 1-2",
            ),
            ("2 TESALONICENSES 1-3", "2 Tesalonicenses 1-3"),
            ("3 JUAN 1-2", "3 Juan 1-2"),
            ("JUDAS 1-4", "Judas 1-4"),
        )

        for source_reading, expected_reading in readings:
            with self.subTest(reading=source_reading):
                header = self._extract_header(
                    f"SEMANA DEL 26 DE ENERO AL 1 DE FEBRERO. {source_reading}"
                )

                self.assertEqual(
                    header["date"],
                    "26 de enero al 1 de febrero",
                )
                self.assertEqual(
                    header["weekly_reading"],
                    expected_reading,
                )

    def test_all_canonical_books_are_accepted_with_chapters(self) -> None:
        self.assertEqual(len(BIBLE_BOOK_NAMES), 66)
        self.assertEqual(len(set(BIBLE_BOOK_NAMES)), 66)

        for book_name in BIBLE_BOOK_NAMES:
            with self.subTest(book=book_name):
                source_reading = f"{book_name} 1-3"
                header = self._extract_header(
                    "SEMANA DEL 2 AL 8 DE FEBRERO. " + source_reading
                )

                self.assertEqual(header["date"], "2 al 8 de febrero")
                self.assertEqual(
                    header["weekly_reading"].casefold(),
                    source_reading.casefold(),
                )

    def test_reading_in_following_paragraph_does_not_change_the_date(self) -> None:
        parser = ProgramParser(Path("unused.docx"))
        document = Document()
        document.add_paragraph("SEMANA DEL 9 AL 15 DE FEBRERO.")
        document.add_paragraph("ISAÍAS 4-6")
        parser.document = document

        headers = parser._extract_week_headers_from_paragraphs()

        self.assertEqual(len(headers), 1)
        self.assertEqual(headers[0]["date"], "9 al 15 de febrero")
        self.assertEqual(headers[0]["weekly_reading"], "Isaías 4-6")


class ScheduleIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temp_directory = tempfile.TemporaryDirectory()
        output_path = Path(cls.temp_directory.name) / "generated.docx"
        cls.weeks = parse_document(SOURCE_DOCUMENT)
        fill_template(TEMPLATE_DOCUMENT, cls.weeks, output_path)
        cls.document = Document(output_path)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp_directory.cleanup()

    def test_first_week_uses_expected_start_times(self) -> None:
        first_copy = _first_form_rows(self.document.tables[0])
        times = [
            row.cells[0].text.strip()
            for row in first_copy
            if row.cells and ":" in row.cells[0].text.strip()
        ]

        self.assertEqual(
            times,
            [
                "7:00",
                "7:05",
                "7:06",
                "7:16",
                "7:26",
                "7:30",
                "7:34",
                "7:39",
                "7:45",
                "7:49",
                "8:04",
                "8:34",
                "8:37",
            ],
        )

    def test_ministry_extra_minute_does_not_change_displayed_duration(self) -> None:
        rows = self.document.tables[0].rows

        self.assertEqual(rows[14].cells[0].text.strip(), "7:30")
        self.assertIn("(3 mins.)", rows[14].cells[1].text)
        self.assertEqual(rows[15].cells[0].text.strip(), "7:34")

    def test_weekly_reading_replaces_template_label(self) -> None:
        header = self.document.tables[0].rows[2].cells[0].text

        self.assertEqual(header, "6 al 12 julio | Jeremías 13-15")
        self.assertNotIn("LECTURA SEMANAL DE LA BIBLIA", header)

    def test_two_participants_use_one_slash(self) -> None:
        participant_text = self.document.tables[0].rows[14].cells[-1].text

        self.assertEqual(participant_text, "Alicia Ejemplo / Beatriz Ejemplo")
        self.assertNotIn("//", participant_text)

    def test_conclusion_participants_are_written_for_every_week(self) -> None:
        rows = [
            row
            for table in self.document.tables
            for row in table.rows
            if "Palabras de conclusión" in row.cells[2].text
        ]

        self.assertEqual(len(rows), len(self.weeks))
        for week, row in zip(self.weeks, rows):
            conclusion = next(
                assignment
                for assignment in week.christian_life
                if assignment.is_conclusion
            )
            self.assertEqual(row.cells[-1].text, conclusion.formatted_participants())

    def test_student_helper_label_is_preserved(self) -> None:
        row_texts = [
            cell.text.strip() for cell in self.document.tables[0].rows[14].cells
        ]

        self.assertIn("Estudiante/Ayudante:", row_texts)

    def test_congregation_name_is_written_in_headers(self) -> None:
        document_text = "\n".join(
            cell.text
            for table in self.document.tables
            for row in table.rows
            for cell in row.cells
        )

        self.assertIn("Congregación de Prueba", document_text)
        self.assertNotIn("[NOMBRE DE LA CONGREGACIÓN]", document_text)

    def test_combined_christian_life_row_has_correct_times(self) -> None:
        rows = [row for table in self.document.tables for row in table.rows]
        first_index = next(
            index
            for index, row in enumerate(rows)
            if "Tema sintético de vida cristiana 1 (10 mins.)" in row.cells[2].text
        )

        self.assertEqual(rows[first_index].cells[0].text.strip(), "7:49")
        self.assertEqual(rows[first_index + 1].cells[0].text.strip(), "7:59")
        self.assertIn("(5 mins.)", rows[first_index + 1].cells[2].text)
        self.assertEqual(rows[first_index + 2].cells[0].text.strip(), "8:04")

    def test_generated_document_has_no_zero_time_markers(self) -> None:
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    self.assertNotIn("0:00", cell.text)

    def test_unused_forms_and_assignment_rows_are_removed(self) -> None:
        headers = [
            cell.text
            for table in self.document.tables
            for row in table.rows
            for cell in row.cells
            if cell.text.strip() == "Presidente:"
        ]
        blank_assignments = [
            cell.text
            for table in self.document.tables
            for row in table.rows
            for cell in row.cells
            if cell.text.strip() in {"7.", "9."}
        ]

        self.assertEqual(len(headers), len(self.weeks))
        self.assertEqual(blank_assignments, [])


class SecondDocumentRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.weeks = parse_document(SECOND_SOURCE_DOCUMENT)
        cls.temp_directory = tempfile.TemporaryDirectory()
        output_path = Path(cls.temp_directory.name) / "generated.docx"
        fill_template(TEMPLATE_DOCUMENT, cls.weeks, output_path)
        cls.document = Document(output_path)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp_directory.cleanup()

    def test_header_and_bible_study_format_variations_are_accepted(self) -> None:
        third_week = self.weeks[2]

        self.assertEqual(third_week.president, "Celia Ejemplo")
        self.assertEqual(third_week.opening_song, "74")
        self.assertEqual(
            [
                assignment.number
                for assignment in third_week.christian_life
                if assignment.is_bible_study
            ],
            [8],
        )

    def test_unnumbered_christian_life_assignments_receive_number(self) -> None:
        for week_index in (4, 5):
            normal_assignments = [
                assignment
                for assignment in self.weeks[week_index].christian_life
                if not assignment.is_bible_study and not assignment.is_conclusion
            ]
            self.assertEqual(
                [
                    (assignment.number, assignment.duration_minutes)
                    for assignment in normal_assignments
                ],
                [(8, 15)],
            )

    def test_ministry_discourse_without_duration_uses_five_minutes(self) -> None:
        discourse = self.weeks[6].ministry[-1]

        self.assertEqual(discourse.number, 7)
        self.assertEqual(discourse.duration_minutes, 5)
        self.assertEqual(discourse.duration_text, "5 mins.")

    def test_variable_ministry_rows_generate_expected_times(self) -> None:
        first_copy = _first_form_rows(self.document.tables[0])
        times = [
            row.cells[0].text.strip()
            for row in first_copy
            if row.cells and ":" in row.cells[0].text.strip()
        ]

        self.assertEqual(
            times,
            [
                "7:00",
                "7:05",
                "7:06",
                "7:16",
                "7:26",
                "7:30",
                "7:34",
                "7:39",
                "7:45",
                "8:00",
                "8:30",
                "8:33",
            ],
        )

    def test_second_generated_document_has_no_zero_time_markers(self) -> None:
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    self.assertNotIn("0:00", cell.text)

    def test_uppercase_source_names_are_normalized(self) -> None:
        first_week = self.weeks[0]

        self.assertEqual(first_week.president, "Mario Prueba")
        self.assertEqual(
            first_week.bible_treasures[0].first_participant_name(),
            "Elena Prueba",
        )

    def test_second_document_includes_weekly_reading(self) -> None:
        header = self.document.tables[0].rows[2].cells[0].text

        self.assertEqual(
            header,
            "7 al 13 de septiembre | Jeremías 32-33",
        )


class ValidationContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.valid_weeks = parse_document(SOURCE_DOCUMENT)

    def setUp(self) -> None:
        self.temp_directory = tempfile.TemporaryDirectory()

    def tearDown(self) -> None:
        self.temp_directory.cleanup()

    def _output_path(self) -> Path:
        return Path(self.temp_directory.name) / "generated.docx"

    def _assert_closed_failure(
        self,
        result: GenerationResult,
        output_path: Path,
        expected_omitted: int,
    ) -> None:
        self.assertFalse(result.succeeded)
        self.assertEqual(result.written_weeks, 0)
        self.assertEqual(result.omitted_weeks, expected_omitted)
        self.assertTrue(result.errors)
        self.assertIsNone(result.output_path)
        self.assertFalse(output_path.exists())

    def test_week_validation_covers_required_invariants(self) -> None:
        cases = {
            "campo requerido": (
                lambda week: setattr(week, "date", ""),
                "fecha",
            ),
            "numeración": (
                lambda week: setattr(week.bible_treasures[1], "number", 1),
                "numer",
            ),
            "participante": (
                lambda week: setattr(
                    week.bible_treasures[0],
                    "participants",
                    [],
                ),
                "participante",
            ),
            "duración": (
                lambda week: setattr(
                    week.bible_treasures[0],
                    "duration_minutes",
                    0,
                ),
                "duración",
            ),
            "canción": (
                lambda week: setattr(week, "opening_song", "desconocida"),
                "canción",
            ),
        }

        for case_name, (mutate, expected_text) in cases.items():
            with self.subTest(case=case_name):
                week = copy.deepcopy(self.valid_weeks[0])
                mutate(week)

                errors = week.validation_errors()

                self.assertFalse(week.is_valid())
                self.assertTrue(
                    any(expected_text in error.lower() for error in errors),
                    errors,
                )

    def test_week_requires_exactly_three_treasure_assignments(self) -> None:
        week = copy.deepcopy(self.valid_weeks[0])
        week.bible_treasures = week.bible_treasures[:1]

        errors = week.validation_errors()

        self.assertFalse(week.is_valid())
        self.assertTrue(
            any("tesoros" in error.lower() and "3" in error for error in errors),
            errors,
        )

    def test_parser_rejects_an_invalid_table_instead_of_returning_it(self) -> None:
        source_path = Path(self.temp_directory.name) / "invalid-source.docx"
        document = Document()
        document.add_table(rows=1, cols=3)
        document.save(source_path)

        with self.assertRaises(GenerationValidationError) as raised:
            parse_document(source_path)

        self.assertTrue(
            any(
                "semana 1" in issue.lower() or "no contiene tablas" in issue.lower()
                for issue in raised.exception.issues
            ),
            raised.exception.issues,
        )

    def test_writer_rejects_invalid_week_without_creating_output(self) -> None:
        week = copy.deepcopy(self.valid_weeks[0])
        week.bible_treasures[0].duration_minutes = 0
        week.bible_treasures[0].duration_text = ""
        output_path = self._output_path()

        result = fill_template(TEMPLATE_DOCUMENT, [week], output_path)

        self.assertTrue(
            any("duración" in issue.lower() for issue in result.errors),
            result.errors,
        )
        self._assert_closed_failure(result, output_path, expected_omitted=1)

    def test_writer_rejects_duration_that_conflicts_with_fixed_template(self) -> None:
        week = copy.deepcopy(self.valid_weeks[0])
        week.bible_treasures[0].duration_minutes = 9
        week.bible_treasures[0].duration_text = "9 mins."
        output_path = self._output_path()

        result = fill_template(TEMPLATE_DOCUMENT, [week], output_path)

        self.assertTrue(
            any("fila fija" in issue.lower() for issue in result.errors),
            result.errors,
        )
        self._assert_closed_failure(result, output_path, expected_omitted=1)

    def test_writer_rejects_more_assignments_than_section_rows(self) -> None:
        week = copy.deepcopy(self.valid_weeks[0])
        last_ministry_assignment = week.ministry[-1]
        week.ministry.extend(
            [
                Assignment(
                    number=7,
                    title="Asignación adicional 1",
                    duration_minutes=5,
                    duration_text="5 mins.",
                    participants=copy.deepcopy(last_ministry_assignment.participants),
                ),
                Assignment(
                    number=8,
                    title="Asignación adicional 2",
                    duration_minutes=5,
                    duration_text="5 mins.",
                    participants=copy.deepcopy(last_ministry_assignment.participants),
                ),
            ]
        )
        for assignment in week.christian_life:
            assignment.number += 2
        output_path = self._output_path()

        result = fill_template(TEMPLATE_DOCUMENT, [week], output_path)

        self.assertTrue(
            any(
                "Seamos Mejores Maestros" in issue and "filas" in issue
                for issue in result.errors
            ),
            result.errors,
        )
        self._assert_closed_failure(result, output_path, expected_omitted=1)

    def test_writer_rejects_more_weeks_than_template_forms(self) -> None:
        template = Document(TEMPLATE_DOCUMENT)
        form_capacity = sum(
            1
            for table in template.tables
            for row in table.rows
            if row.cells and "[FECHA]" in row.cells[0].text
        )
        weeks = [
            copy.deepcopy(self.valid_weeks[index % len(self.valid_weeks)])
            for index in range(form_capacity + 1)
        ]
        output_path = self._output_path()

        result = fill_template(TEMPLATE_DOCUMENT, weeks, output_path)

        self.assertTrue(
            any(
                "plantilla" in issue.lower()
                and str(form_capacity) in issue
                and str(form_capacity + 1) in issue
                for issue in result.errors
            ),
            result.errors,
        )
        self._assert_closed_failure(
            result,
            output_path,
            expected_omitted=form_capacity + 1,
        )

    def test_writer_rejects_template_without_tables(self) -> None:
        template_path = Path(self.temp_directory.name) / "empty-template.docx"
        Document().save(template_path)
        output_path = self._output_path()

        result = fill_template(template_path, [self.valid_weeks[0]], output_path)

        self.assertTrue(
            any("formularios" in issue for issue in result.errors),
            result.errors,
        )
        self._assert_closed_failure(result, output_path, expected_omitted=1)

    def test_writer_rejects_tables_without_recognizable_forms(self) -> None:
        template_path = Path(self.temp_directory.name) / "no-forms-template.docx"
        template = Document()
        template.add_table(rows=1, cols=1).cell(0, 0).text = "Sin formulario"
        template.save(template_path)
        output_path = self._output_path()

        result = fill_template(template_path, [self.valid_weeks[0]], output_path)

        self.assertTrue(
            any("formularios" in issue for issue in result.errors),
            result.errors,
        )
        self._assert_closed_failure(result, output_path, expected_omitted=1)

    def test_deleted_placeholder_is_not_mistaken_for_written_value(self) -> None:
        output_path = self._output_path()

        def erase_participants(writer, cell, _assignment) -> None:
            writer._empty_cell(cell)

        with patch.object(
            TemplateWriter,
            "_write_participants_in_cell",
            new=erase_participants,
        ):
            result = fill_template(
                TEMPLATE_DOCUMENT,
                [self.valid_weeks[0]],
                output_path,
            )

        self.assertTrue(
            any(
                "no se escribieron los participantes" in issue
                for issue in result.errors
            ),
            result.errors,
        )
        self._assert_closed_failure(result, output_path, expected_omitted=1)

    def test_writer_rejects_marker_that_cannot_be_cleared_from_unused_copy(
        self,
    ) -> None:
        output_path = self._output_path()
        original_remove_text = TemplateWriter._remove_text

        def leave_title_marker(writer, cell, text: str) -> None:
            if text != "[Título]":
                original_remove_text(writer, cell, text)

        with patch.object(
            TemplateWriter,
            "_remove_text",
            new=leave_title_marker,
        ):
            result = fill_template(
                TEMPLATE_DOCUMENT,
                [self.valid_weeks[0]],
                output_path,
            )

        self.assertIn("[Título]", result.errors[0])
        self._assert_closed_failure(result, output_path, expected_omitted=1)


class GenerationResultTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.valid_weeks = parse_document(SOURCE_DOCUMENT)

    def setUp(self) -> None:
        self.temp_directory = tempfile.TemporaryDirectory()
        self.output_path = Path(self.temp_directory.name) / "generated.docx"

    def tearDown(self) -> None:
        self.temp_directory.cleanup()

    def test_writer_success_reports_only_weeks_actually_written(self) -> None:
        weeks = self.valid_weeks[:2]

        result = fill_template(TEMPLATE_DOCUMENT, weeks, self.output_path)

        self.assertIsInstance(result, GenerationResult)
        self.assertTrue(result.succeeded)
        self.assertEqual(result.written_weeks, 2)
        self.assertEqual(result.omitted_weeks, 0)
        self.assertEqual(result.errors, ())
        self.assertEqual(result.output_path, self.output_path)
        self.assertTrue(self.output_path.is_file())
        self.assertGreater(self.output_path.stat().st_size, 0)

    def test_horizontal_merges_preserve_main_participants_and_clear_auxiliary_cells(
        self,
    ) -> None:
        template_path = Path(self.temp_directory.name) / "horizontal-merges.docx"
        document = Document(TEMPLATE_DOCUMENT)
        table = document.tables[0]
        week = self.valid_weeks[0]
        assignments = (
            (9, 2, week.bible_treasures[0]),
            (14, 1, week.ministry[0]),
            (21, 2, week.christian_life[0]),
            (
                23,
                2,
                next(
                    assignment
                    for assignment in week.christian_life
                    if assignment.is_bible_study
                ),
            ),
        )
        for row_index, title_column, assignment in assignments:
            row = table.rows[row_index]
            row.cells[title_column].merge(row.cells[5])
            auxiliary = row.cells[8].merge(row.cells[9])
            auxiliary.text = row.cells[13].text
            participant = row.cells[11].merge(row.cells[13])
            participant.paragraphs[0].runs[0].bold = True
        document.save(template_path)
        original_template = template_path.read_bytes()
        writer = TemplateWriter(template_path)

        with patch.object(
            writer, "_write_number_in_cell", wraps=writer._write_number_in_cell
        ) as write_number:
            result = writer.fill([week], self.output_path)

        self.assertTrue(result.succeeded, result.errors)
        written_cells = [call.args[0]._tc for call in write_number.call_args_list]
        self.assertEqual(len(written_cells), len(set(written_cells)))
        output = Document(self.output_path)
        for row_index, title_column, assignment in assignments:
            row = next(
                row
                for row in output.tables[0].rows
                if assignment.title in row.cells[title_column].text
            )
            participant = row.cells[13]
            self.assertIs(row.cells[11]._tc, participant._tc)
            self.assertEqual(participant.text, assignment.formatted_participants())
            self.assertTrue(participant.paragraphs[0].runs[0].bold)
            self.assertEqual(row.cells[8].text, "")
            self.assertEqual(row.cells[9].text, "")
            self.assertIn(assignment.title, row.cells[title_column].text)
            self.assertIn(assignment.duration_text, row.cells[title_column].text)
        self.assertEqual(
            output.tables[0].rows[14].cells[10].text, "Estudiante/Ayudante:"
        )
        self.assertEqual(template_path.read_bytes(), original_template)

    def test_vertical_and_rectangular_merges_write_shared_participants_once(
        self,
    ) -> None:
        for start_column in (13, 11):
            with self.subTest(start_column=start_column):
                template_path = Path(self.temp_directory.name) / "vertical-merges.docx"
                document = Document(TEMPLATE_DOCUMENT)
                table = document.tables[0]
                participant = table.cell(14, start_column).merge(table.cell(16, 13))
                participant.text = "[Nombre/Nombre]"
                auxiliary = table.cell(14, 8).merge(table.cell(16, 8))
                auxiliary.text = "[Nombre/Nombre]"
                document.save(template_path)
                writer = TemplateWriter(template_path)

                with (
                    patch.object(
                        writer,
                        "_write_participants_in_cell",
                        wraps=writer._write_participants_in_cell,
                    ) as write_participants,
                    patch.object(
                        writer, "_empty_cell", wraps=writer._empty_cell
                    ) as empty,
                ):
                    result = writer.fill([self.valid_weeks[0]], self.output_path)

                self.assertTrue(result.succeeded, result.errors)
                cleared_cells = [call.args[0]._tc for call in empty.call_args_list]
                self.assertEqual(len(cleared_cells), len(set(cleared_cells)))
                output = Document(self.output_path)
                shared_cell = output.tables[0].cell(14, 13)
                for row_index in (14, 15, 16):
                    row = output.tables[0].rows[row_index]
                    self.assertIs(row.cells[13]._tc, shared_cell._tc)
                    self.assertEqual(
                        row.cells[13].text,
                        self.valid_weeks[0]
                        .ministry[row_index - 14]
                        .formatted_participants(),
                    )
                    self.assertEqual(row.cells[8].text, "")
                    self.assertEqual(row.cells[10].text, "Estudiante/Ayudante:")
                shared_writes = [
                    call
                    for call in write_participants.call_args_list
                    if call.args[0]._tc.getparent()
                    is writer.document.tables[0].rows[14]._tr
                ]
                self.assertEqual(len(shared_writes), 1)

    def test_vertical_continuations_do_not_create_rows_or_erase_schedule_time(
        self,
    ) -> None:
        template_path = Path(self.temp_directory.name) / "vertical-continuations.docx"
        document = Document(TEMPLATE_DOCUMENT)
        table = document.tables[0]
        table.cell(2, 0).merge(table.cell(3, 0))
        for column in (0, 2, 13):
            table.cell(11, column).merge(table.cell(12, column))
        document.save(template_path)
        writer = TemplateWriter(template_path)
        writer.document = Document(template_path)

        copies = writer._discover_template_copies()

        self.assertEqual(len(copies), 10)
        self.assertEqual(
            [row.row_index for row in copies[0].structure["bible_treasures_rows"]],
            [9, 10, 11],
        )
        result = writer.fill([self.valid_weeks[0]], self.output_path)

        self.assertTrue(result.succeeded, result.errors)
        output = Document(self.output_path)
        self.assertEqual(output.tables[0].cell(11, 0).text, "7:26")
        self.assertEqual(output.tables[0].cell(12, 0).text, "7:26")
        self.assertEqual(
            output.tables[0].cell(12, 13).text,
            self.valid_weeks[0].bible_treasures[2].formatted_participants(),
        )

    def test_empty_time_cells_are_filled_and_preserve_format(self) -> None:
        template_path = Path(self.temp_directory.name) / "empty-times.docx"
        expected_times = [
            "7:00",
            "7:05",
            "7:06",
            "7:16",
            "7:26",
            "7:30",
            "7:34",
            "7:39",
            "7:45",
            "7:49",
            "8:04",
            "8:34",
            "8:37",
        ]
        for layout in ("without_runs", "formatted_run", "whitespace"):
            with self.subTest(layout=layout):
                document = Document(TEMPLATE_DOCUMENT)
                table = document.tables[0]
                for row_index in (5, 6, 9, 10, 11, 14, 15, 16, 20, 21, 23, 24, 25):
                    paragraph = table.cell(row_index, 0).paragraphs[0]
                    paragraph.clear()
                    if layout != "without_runs":
                        run = paragraph.add_run(
                            " \t " if layout == "whitespace" else ""
                        )
                        run.bold = True
                        run.italic = True
                document.save(template_path)
                original_template = template_path.read_bytes()

                result = fill_template(
                    template_path, [self.valid_weeks[0]], self.output_path
                )

                self.assertTrue(result.succeeded, result.errors)
                output = Document(self.output_path)
                time_cells = [
                    row.cells[0]
                    for row in _first_form_rows(output.tables[0])
                    if ":" in row.cells[0].text
                ]
                self.assertEqual(
                    [cell.text.strip() for cell in time_cells], expected_times
                )
                if layout != "without_runs":
                    for cell in time_cells:
                        self.assertTrue(cell.paragraphs[0].runs[0].bold)
                        self.assertTrue(cell.paragraphs[0].runs[0].italic)
                self.assertEqual(template_path.read_bytes(), original_template)

    def test_obsolete_time_in_empty_row_is_cleared(self) -> None:
        template_path = Path(self.temp_directory.name) / "obsolete-time.docx"
        document = Document(TEMPLATE_DOCUMENT)
        document.tables[0].cell(7, 0).text = "9:59"
        document.save(template_path)
        original_template = template_path.read_bytes()

        result = fill_template(template_path, [self.valid_weeks[0]], self.output_path)

        self.assertTrue(result.succeeded, result.errors)
        output = Document(self.output_path)
        self.assertEqual(output.tables[0].cell(7, 0).text, "")
        self.assertEqual(
            [
                row.cells[0].text.strip()
                for row in _first_form_rows(output.tables[0])
                if ":" in row.cells[0].text
            ],
            [
                "7:00",
                "7:05",
                "7:06",
                "7:16",
                "7:26",
                "7:30",
                "7:34",
                "7:39",
                "7:45",
                "7:49",
                "8:04",
                "8:34",
                "8:37",
            ],
        )
        self.assertEqual(template_path.read_bytes(), original_template)

    def test_missing_time_column_rejects_generation_and_preserves_output(self) -> None:
        template_path = Path(self.temp_directory.name) / "missing-time-column.docx"
        previous_content = b"salida anterior"
        for row_index in (5, 9):
            for existing_output in (False, True):
                with self.subTest(row_index=row_index, existing_output=existing_output):
                    document = Document(TEMPLATE_DOCUMENT)
                    row = document.tables[0].rows[row_index]
                    row._tr.remove(row.cells[0]._tc)
                    grid_before = OxmlElement("w:gridBefore")
                    grid_before.set(qn("w:val"), "1")
                    row._tr.get_or_add_trPr().append(grid_before)
                    document.save(template_path)
                    original_template = template_path.read_bytes()
                    self.output_path.unlink(missing_ok=True)
                    if existing_output:
                        self.output_path.write_bytes(previous_content)
                    files_before = set(self.output_path.parent.iterdir())

                    with patch(
                        "meeting_generator.template_writer.os.replace",
                        wraps=os.replace,
                    ) as replace:
                        result = fill_template(
                            template_path, [self.valid_weeks[0]], self.output_path
                        )

                    self.assertFalse(result.succeeded)
                    self.assertEqual(result.written_weeks, 0)
                    self.assertEqual(result.omitted_weeks, 1)
                    self.assertIsNone(result.output_path)
                    self.assertIn(
                        f"semana 1, tabla 1, fila {row_index + 1}: "
                        "falta la celda horaria",
                        result.errors,
                    )
                    replace.assert_not_called()
                    self.assertEqual(template_path.read_bytes(), original_template)
                    self.assertEqual(
                        set(self.output_path.parent.iterdir()), files_before
                    )
                    if existing_output:
                        self.assertEqual(
                            self.output_path.read_bytes(), previous_content
                        )
                    else:
                        self.assertFalse(self.output_path.exists())

    def test_time_merge_from_retained_empty_rows_preserves_schedule(self) -> None:
        template_path = Path(self.temp_directory.name) / "retained-time-origin.docx"
        for marker in ("", "7:00"):
            with self.subTest(marker=marker):
                document = Document(TEMPLATE_DOCUMENT)
                table = document.tables[0]
                table.cell(3, 0).merge(table.cell(5, 0)).text = marker
                document.save(template_path)
                original_template = template_path.read_bytes()

                result = fill_template(
                    template_path, [self.valid_weeks[0]], self.output_path
                )

                self.assertTrue(result.succeeded, result.errors)
                output = Document(self.output_path)
                table = output.tables[0]
                for row_index in (3, 4, 5):
                    self.assertEqual(table.cell(row_index, 0).text, "7:00")
                    self.assertIs(table.cell(row_index, 0)._tc, table.cell(3, 0)._tc)
                self.assertEqual(
                    table.cell(5, 2).text,
                    f"Canción {self.valid_weeks[0].opening_song}",
                )
                times = [
                    row.cells[0].text.strip()
                    for row in _first_form_rows(table)
                    if ":" in row.cells[0].text
                    and any(cell.text.strip() for cell in row.cells[1:])
                ]
                self.assertEqual(
                    times,
                    [
                        "7:00",
                        "7:05",
                        "7:06",
                        "7:16",
                        "7:26",
                        "7:30",
                        "7:34",
                        "7:39",
                        "7:45",
                        "7:49",
                        "8:04",
                        "8:34",
                        "8:37",
                    ],
                )
                self.assertEqual(template_path.read_bytes(), original_template)

    def test_time_merge_with_section_heading_rejects_generation(self) -> None:
        template_path = Path(self.temp_directory.name) / "section-time-merge.docx"
        previous_content = b"salida anterior"
        for layout in ("merged_paragraphs", "single_paragraph"):
            for existing_output in (False, True):
                with self.subTest(layout=layout, existing_output=existing_output):
                    document = Document(TEMPLATE_DOCUMENT)
                    table = document.tables[0]
                    cell = table.cell(8, 0).merge(table.cell(9, 0))
                    if layout == "single_paragraph":
                        cell.text = "TESOROS DE LA BIBLIA"
                    document.save(template_path)
                    original_template = template_path.read_bytes()
                    self.output_path.unlink(missing_ok=True)
                    if existing_output:
                        self.output_path.write_bytes(previous_content)
                    files_before = set(self.output_path.parent.iterdir())

                    with patch(
                        "meeting_generator.template_writer.os.replace",
                        wraps=os.replace,
                    ) as replace:
                        result = fill_template(
                            template_path, [self.valid_weeks[0]], self.output_path
                        )

                    self.assertFalse(result.succeeded)
                    self.assertEqual(result.written_weeks, 0)
                    self.assertEqual(result.omitted_weeks, 1)
                    self.assertIsNone(result.output_path)
                    self.assertIn(
                        "semana 1, tabla 1, fila 10: no se escribió la hora esperada 7:06",
                        result.errors,
                    )
                    replace.assert_not_called()
                    self.assertEqual(template_path.read_bytes(), original_template)
                    self.assertEqual(
                        set(self.output_path.parent.iterdir()), files_before
                    )
                    if existing_output:
                        self.assertEqual(
                            self.output_path.read_bytes(), previous_content
                        )
                    else:
                        self.assertFalse(self.output_path.exists())

    def test_time_merge_with_introduction_content_rejects_generation(self) -> None:
        template_path = Path(self.temp_directory.name) / "introduction-time-merge.docx"
        previous_content = b"salida anterior"
        document = Document(TEMPLATE_DOCUMENT)
        table = document.tables[0]
        cell = table.cell(6, 0).merge(table.cell(6, 2))
        cell.text = "0:00 Palabras de introducción (1 min.)"
        document.save(template_path)
        original_template = template_path.read_bytes()
        for existing_output in (False, True):
            with self.subTest(existing_output=existing_output):
                self.output_path.unlink(missing_ok=True)
                if existing_output:
                    self.output_path.write_bytes(previous_content)
                files_before = set(self.output_path.parent.iterdir())

                with patch(
                    "meeting_generator.template_writer.os.replace",
                    wraps=os.replace,
                ) as replace:
                    result = fill_template(
                        template_path, [self.valid_weeks[0]], self.output_path
                    )

                self.assertFalse(result.succeeded)
                self.assertEqual(result.written_weeks, 0)
                self.assertEqual(result.omitted_weeks, 1)
                self.assertIsNone(result.output_path)
                self.assertIn(
                    "semana 1, tabla 1, fila 7: no se escribió la hora esperada 7:05",
                    result.errors,
                )
                replace.assert_not_called()
                self.assertEqual(template_path.read_bytes(), original_template)
                self.assertEqual(set(self.output_path.parent.iterdir()), files_before)
                if existing_output:
                    self.assertEqual(self.output_path.read_bytes(), previous_content)
                else:
                    self.assertFalse(self.output_path.exists())

    def test_incompatible_time_merges_reject_generation_and_preserve_output(
        self,
    ) -> None:
        template_path = Path(self.temp_directory.name) / "merged-times.docx"
        previous_content = b"salida anterior"
        for marker in ("0:00", "", None):
            for existing_output in (False, True):
                with self.subTest(marker=marker, existing_output=existing_output):
                    document = Document(TEMPLATE_DOCUMENT)
                    cell = (
                        document.tables[0]
                        .cell(9, 0)
                        .merge(document.tables[0].cell(10, 0))
                    )
                    if marker is not None:
                        cell.text = marker
                    document.save(template_path)
                    original_template = template_path.read_bytes()
                    self.output_path.unlink(missing_ok=True)
                    if existing_output:
                        self.output_path.write_bytes(previous_content)
                    files_before = set(self.output_path.parent.iterdir())

                    with patch.object(
                        TemplateWriter, "_save_document_atomically"
                    ) as save:
                        result = fill_template(
                            template_path, [self.valid_weeks[0]], self.output_path
                        )

                    self.assertFalse(result.succeeded)
                    self.assertEqual(result.written_weeks, 0)
                    self.assertEqual(result.omitted_weeks, 1)
                    self.assertIsNone(result.output_path)
                    self.assertTrue(
                        any(
                            "semana 1, tabla 1, fila 11" in error
                            and "celda horaria fusionada con la fila 10" in error
                            and "7:06 y 7:16" in error
                            for error in result.errors
                        ),
                        result.errors,
                    )
                    save.assert_not_called()
                    self.assertEqual(template_path.read_bytes(), original_template)
                    self.assertEqual(
                        set(self.output_path.parent.iterdir()), files_before
                    )
                    if existing_output:
                        self.assertEqual(
                            self.output_path.read_bytes(), previous_content
                        )
                    else:
                        self.assertFalse(self.output_path.exists())

    def test_missing_or_incorrect_written_times_preserve_previous_output(self) -> None:
        previous_content = b"salida anterior"
        write_time = TemplateWriter._write_time_in_row
        for actual_time in ("", "9:59"):
            with self.subTest(actual_time=actual_time):
                self.output_path.write_bytes(previous_content)

                def write_wrong_time(writer, table, row_index, total_minutes):
                    write_time(writer, table, row_index, total_minutes)
                    if row_index == 9:
                        table.cell(row_index, 0).text = actual_time

                with (
                    patch.object(
                        TemplateWriter, "_write_time_in_row", write_wrong_time
                    ),
                    patch.object(TemplateWriter, "_save_document_atomically") as save,
                ):
                    result = fill_template(
                        TEMPLATE_DOCUMENT, [self.valid_weeks[0]], self.output_path
                    )

                self.assertFalse(result.succeeded)
                self.assertEqual(result.written_weeks, 0)
                self.assertEqual(result.omitted_weeks, 1)
                self.assertIsNone(result.output_path)
                self.assertIn(
                    "semana 1, tabla 1, fila 10: no se escribió la hora esperada 7:06",
                    result.errors,
                )
                save.assert_not_called()
                self.assertEqual(self.output_path.read_bytes(), previous_content)

    def test_time_merge_origin_in_unused_row_rejects_generation(self) -> None:
        template_path = Path(self.temp_directory.name) / "unused-time-origin.docx"
        previous_content = b"salida anterior"
        for marker in ("0:00", "", "8:04", None):
            for existing_output in (False, True):
                with self.subTest(marker=marker, existing_output=existing_output):
                    document = Document(TEMPLATE_DOCUMENT)
                    table = document.tables[0]
                    cell = table.cell(22, 0).merge(table.cell(23, 0))
                    if marker is not None:
                        cell.text = marker
                    document.save(template_path)
                    original_template = template_path.read_bytes()
                    self.output_path.unlink(missing_ok=True)
                    if existing_output:
                        self.output_path.write_bytes(previous_content)
                    files_before = set(self.output_path.parent.iterdir())

                    with patch.object(
                        TemplateWriter,
                        "_save_document_atomically",
                        return_value=self.output_path,
                    ) as save:
                        result = fill_template(
                            template_path, [self.valid_weeks[0]], self.output_path
                        )

                    self.assertFalse(result.succeeded)
                    self.assertEqual(result.written_weeks, 0)
                    self.assertEqual(result.omitted_weeks, 1)
                    self.assertIsNone(result.output_path)
                    self.assertTrue(
                        any(
                            "semana 1, tabla 1, fila 24" in error
                            and "hora esperada 8:04" in error
                            for error in result.errors
                        ),
                        result.errors,
                    )
                    save.assert_not_called()
                    self.assertEqual(template_path.read_bytes(), original_template)
                    self.assertEqual(
                        set(self.output_path.parent.iterdir()), files_before
                    )
                    if existing_output:
                        self.assertEqual(
                            self.output_path.read_bytes(), previous_content
                        )
                    else:
                        self.assertFalse(self.output_path.exists())

    def test_times_changed_during_cleanup_reject_generation_before_publication(
        self,
    ) -> None:
        previous_content = b"salida anterior"
        remove_unused_structure = TemplateWriter._remove_unused_structure
        for actual_time in ("", "7:49"):
            for existing_output in (False, True):
                with self.subTest(
                    actual_time=actual_time, existing_output=existing_output
                ):
                    self.output_path.unlink(missing_ok=True)
                    if existing_output:
                        self.output_path.write_bytes(previous_content)
                    files_before = set(self.output_path.parent.iterdir())

                    def remove_structure_and_change_time(writer, copies, weeks):
                        remove_unused_structure(writer, copies, weeks)
                        row = next(
                            row
                            for row in writer.document.tables[0].rows
                            if any(
                                "Estudio bíblico de la congregación" in cell.text
                                for cell in row.cells
                            )
                        )
                        row.cells[0].text = actual_time

                    with (
                        patch.object(
                            TemplateWriter,
                            "_remove_unused_structure",
                            remove_structure_and_change_time,
                        ),
                        patch(
                            "meeting_generator.template_writer.os.replace",
                            wraps=os.replace,
                        ) as replace,
                    ):
                        result = fill_template(
                            TEMPLATE_DOCUMENT, [self.valid_weeks[0]], self.output_path
                        )

                    self.assertFalse(result.succeeded)
                    self.assertEqual(result.written_weeks, 0)
                    self.assertEqual(result.omitted_weeks, 1)
                    self.assertIsNone(result.output_path)
                    self.assertTrue(
                        any(
                            "semana 1, tabla 1" in error
                            and "hora" in error
                            and "8:04" in error
                            for error in result.errors
                        ),
                        result.errors,
                    )
                    replace.assert_not_called()
                    self.assertEqual(
                        set(self.output_path.parent.iterdir()), files_before
                    )
                    if existing_output:
                        self.assertEqual(
                            self.output_path.read_bytes(), previous_content
                        )
                    else:
                        self.assertFalse(self.output_path.exists())

    def test_times_changed_in_serialized_document_reject_publication(self) -> None:
        previous_content = b"salida anterior"
        save_document = DocxDocument.save
        for actual_time in ("", "7:49"):
            for existing_output in (False, True):
                with self.subTest(
                    actual_time=actual_time, existing_output=existing_output
                ):
                    self.output_path.unlink(missing_ok=True)
                    if existing_output:
                        self.output_path.write_bytes(previous_content)
                    files_before = set(self.output_path.parent.iterdir())
                    temporary_paths = []

                    def save_with_changed_time(document, file):
                        save_document(document, file)
                        temporary_paths.append(Path(file))
                        serialized_document = Document(file)
                        row = next(
                            row
                            for row in serialized_document.tables[0].rows
                            if any(
                                "Estudio bíblico de la congregación" in cell.text
                                for cell in row.cells
                            )
                        )
                        row.cells[0].text = actual_time
                        save_document(serialized_document, file)

                    with (
                        patch.object(DocxDocument, "save", save_with_changed_time),
                        patch(
                            "meeting_generator.template_writer.os.replace",
                            wraps=os.replace,
                        ) as replace,
                    ):
                        result = fill_template(
                            TEMPLATE_DOCUMENT, [self.valid_weeks[0]], self.output_path
                        )

                    self.assertFalse(result.succeeded)
                    self.assertEqual(result.written_weeks, 0)
                    self.assertEqual(result.omitted_weeks, 1)
                    self.assertIsNone(result.output_path)
                    self.assertTrue(
                        any(
                            "semana 1, tabla 1" in error
                            and "hora" in error
                            and "8:04" in error
                            for error in result.errors
                        ),
                        result.errors,
                    )
                    replace.assert_not_called()
                    self.assertEqual(len(temporary_paths), 1)
                    self.assertEqual(temporary_paths[0].parent, self.output_path.parent)
                    self.assertFalse(temporary_paths[0].exists())
                    self.assertEqual(
                        set(self.output_path.parent.iterdir()), files_before
                    )
                    if existing_output:
                        self.assertEqual(
                            self.output_path.read_bytes(), previous_content
                        )
                    else:
                        self.assertFalse(self.output_path.exists())

    def test_missing_time_column_in_serialized_document_rejects_publication(
        self,
    ) -> None:
        previous_content = b"salida anterior"
        original_template = TEMPLATE_DOCUMENT.read_bytes()
        save_document = DocxDocument.save
        for existing_output in (False, True):
            with self.subTest(existing_output=existing_output):
                self.output_path.unlink(missing_ok=True)
                if existing_output:
                    self.output_path.write_bytes(previous_content)
                files_before = set(self.output_path.parent.iterdir())
                temporary_paths = []

                def save_without_time_column(document, file):
                    save_document(document, file)
                    temporary_paths.append(Path(file))
                    serialized_document = Document(file)
                    row = serialized_document.tables[0].rows[5]
                    row._tr.remove(row.cells[0]._tc)
                    grid_before = OxmlElement("w:gridBefore")
                    grid_before.set(qn("w:val"), "1")
                    row._tr.get_or_add_trPr().append(grid_before)
                    row.cells[0].text = "7:00"
                    save_document(serialized_document, file)

                with (
                    patch.object(DocxDocument, "save", save_without_time_column),
                    patch(
                        "meeting_generator.template_writer.os.replace",
                        wraps=os.replace,
                    ) as replace,
                ):
                    result = fill_template(
                        TEMPLATE_DOCUMENT, [self.valid_weeks[0]], self.output_path
                    )

                self.assertFalse(result.succeeded)
                self.assertEqual(result.written_weeks, 0)
                self.assertEqual(result.omitted_weeks, 1)
                self.assertIsNone(result.output_path)
                self.assertIn(
                    "semana 1, tabla 1, fila 6: falta la celda horaria", result.errors
                )
                replace.assert_not_called()
                self.assertEqual(len(temporary_paths), 1)
                self.assertEqual(temporary_paths[0].parent, self.output_path.parent)
                self.assertFalse(temporary_paths[0].exists())
                self.assertEqual(TEMPLATE_DOCUMENT.read_bytes(), original_template)
                self.assertEqual(set(self.output_path.parent.iterdir()), files_before)
                if existing_output:
                    self.assertEqual(self.output_path.read_bytes(), previous_content)
                else:
                    self.assertFalse(self.output_path.exists())

    def test_incompatible_vertical_participants_preserve_previous_output(self) -> None:
        template_path = Path(self.temp_directory.name) / "incompatible-merge.docx"
        document = Document(TEMPLATE_DOCUMENT)
        table = document.tables[0]
        table.cell(14, 13).merge(table.cell(15, 13)).text = "[Nombre/Nombre]"
        document.save(template_path)
        previous_content = b"salida anterior"
        for case in ("otro nombre", "solo el estudiante"):
            with self.subTest(case=case):
                week = copy.deepcopy(self.valid_weeks[0])
                if case == "otro nombre":
                    week.ministry[1].participants[0].name = "Amelia Ejemplo"
                else:
                    week.ministry[1].participants = week.ministry[1].participants[:1]
                self.output_path.write_bytes(previous_content)

                result = fill_template(template_path, [week], self.output_path)

                self.assertFalse(result.succeeded)
                self.assertEqual(result.written_weeks, 0)
                self.assertIsNone(result.output_path)
                self.assertTrue(
                    any(
                        "fila 16" in error
                        and "participantes distintos" in error
                        and "punto 5" in error
                        for error in result.errors
                    ),
                    result.errors,
                )
                self.assertEqual(self.output_path.read_bytes(), previous_content)

    def test_full_workflow_reports_exact_success_counts_and_real_path(self) -> None:
        result = generate_document(
            SOURCE_DOCUMENT,
            TEMPLATE_DOCUMENT,
            self.output_path,
        )

        self.assertTrue(result.succeeded)
        self.assertEqual(result.written_weeks, len(self.valid_weeks))
        self.assertEqual(result.omitted_weeks, 0)
        self.assertEqual(result.errors, ())
        self.assertEqual(result.output_path, self.output_path)
        self.assertTrue(self.output_path.is_file())

    def test_conclusion_writes_its_participant_with_blank_or_merged_markers(
        self,
    ) -> None:
        template_path = Path(self.temp_directory.name) / "conclusion-template.docx"
        week = copy.deepcopy(self.valid_weeks[0])
        conclusion = next(
            assignment for assignment in week.christian_life if assignment.is_conclusion
        )
        conclusion.participants[0].name = "Marcos Ejemplo (h)"
        self.assertNotEqual(conclusion.formatted_participants(), week.president)
        for layout, marker in (
            ("empty", ""),
            ("short_row", ""),
            ("marker", "[Nombre]"),
            ("horizontal", "[Nombre/Nombre]"),
            ("vertical", "[Nombre]"),
        ):
            with self.subTest(layout=layout):
                document = Document(TEMPLATE_DOCUMENT)
                table = document.tables[0]
                cell = table.cell(24, 13)
                if layout == "short_row":
                    table.cell(24, 0).merge(table.cell(24, 1))
                    table.cell(24, 2).merge(table.cell(24, 11))
                    row = table.rows[24]
                    row._tr.remove(row.cells[13]._tc)
                    cell = table.rows[24].cells[-1]
                elif layout == "horizontal":
                    cell = table.cell(24, 11).merge(cell)
                elif layout == "vertical":
                    table.cell(25, 12).text = "[Nombre]"
                    table.cell(25, 13).text = ""
                    cell = cell.merge(table.cell(25, 13))
                cell.text = ""
                paragraph = cell.paragraphs[0]
                paragraph.clear()
                run = paragraph.add_run(marker)
                run.bold = True
                document.save(template_path)
                original_template = template_path.read_bytes()

                result = fill_template(template_path, [week], self.output_path)

                self.assertTrue(result.succeeded, result.errors)
                output = Document(self.output_path)
                row = next(
                    row
                    for row in output.tables[0].rows
                    if "Palabras de conclusión" in row.cells[2].text
                )
                self.assertEqual(row.cells[-1].text, "Marcos Ejemplo (h)")
                self.assertTrue(row.cells[-1].paragraphs[0].runs[0].bold)
                self.assertEqual(row.cells[2].text, "Palabras de conclusión (3 mins.)")
                self.assertEqual(row.cells[0].text, "8:34")
                closing_row = next(
                    row
                    for row in output.tables[0].rows
                    if "Oración:" in row.cells[9].text
                    and f"Canción {week.closing_song}" in row.cells[2].text
                )
                self.assertIn(
                    week.closing_prayer, [cell.text for cell in closing_row.cells]
                )
                self.assertEqual(template_path.read_bytes(), original_template)

    def test_missing_conclusion_in_output_never_creates_or_replaces_file(self) -> None:
        original_write = TemplateWriter._write_participants_in_cell

        def skip_conclusion(writer, cell, assignment):
            if not assignment.is_conclusion:
                original_write(writer, cell, assignment)

        for existing_output in (False, True):
            with self.subTest(existing_output=existing_output):
                output_path = (
                    Path(self.temp_directory.name)
                    / f"conclusion-{existing_output}.docx"
                )
                previous_content = b"salida anterior"
                if existing_output:
                    output_path.write_bytes(previous_content)

                with patch.object(
                    TemplateWriter, "_write_participants_in_cell", new=skip_conclusion
                ):
                    result = fill_template(
                        TEMPLATE_DOCUMENT, [self.valid_weeks[0]], output_path
                    )

                self.assertFalse(result.succeeded)
                self.assertEqual(result.written_weeks, 0)
                self.assertIsNone(result.output_path)
                self.assertTrue(
                    any(
                        "participantes de la conclusión" in error
                        for error in result.errors
                    ),
                    result.errors,
                )
                if existing_output:
                    self.assertEqual(output_path.read_bytes(), previous_content)
                else:
                    self.assertFalse(output_path.exists())

    def test_conclusion_rejects_unsafe_cells_or_missing_participant(self) -> None:
        template_path = Path(self.temp_directory.name) / "invalid-conclusion.docx"
        previous_content = b"salida anterior"
        for case in (
            "occupied",
            "merged_with_title",
            "shared_with_study",
            "missing_name",
        ):
            with self.subTest(case=case):
                document = Document(TEMPLATE_DOCUMENT)
                table = document.tables[0]
                week = copy.deepcopy(self.valid_weeks[0])
                if case == "occupied":
                    table.cell(24, 13).text = "Texto que debe conservarse"
                elif case == "merged_with_title":
                    table.cell(24, 2).merge(table.cell(24, 13))
                elif case == "shared_with_study":
                    table.cell(23, 13).merge(table.cell(24, 13))
                else:
                    conclusion = next(
                        assignment
                        for assignment in week.christian_life
                        if assignment.is_conclusion
                    )
                    conclusion.participants = []
                document.save(template_path)
                original_template = template_path.read_bytes()
                self.output_path.write_bytes(previous_content)

                result = fill_template(template_path, [week], self.output_path)

                self.assertFalse(result.succeeded)
                self.assertEqual(result.written_weeks, 0)
                self.assertTrue(result.errors)
                self.assertEqual(self.output_path.read_bytes(), previous_content)
                self.assertEqual(template_path.read_bytes(), original_template)

    def test_full_workflow_uses_corrections_only_for_requested_document(self) -> None:
        original_source = SOURCE_DOCUMENT.read_bytes()
        corrections = (
            NameCorrection(
                original_name="Alicia Ejemplo",
                corrected_name="Amelia Ejemplo",
                reason="Nombre confirmado en la referencia sintética de prueba",
            ),
        )

        result = generate_document(
            SOURCE_DOCUMENT,
            TEMPLATE_DOCUMENT,
            self.output_path,
            name_corrections=corrections,
        )

        self.assertTrue(result.succeeded, result.errors)
        document = Document(self.output_path)
        self.assertEqual(
            document.tables[0].rows[14].cells[13].text,
            "Amelia Ejemplo / Beatriz Ejemplo",
        )
        self.assertEqual(SOURCE_DOCUMENT.read_bytes(), original_source)
        self.assertEqual(
            parse_document(SOURCE_DOCUMENT)[0].ministry[0].first_participant_name(),
            "Alicia Ejemplo",
        )

    def test_duplicate_name_corrections_preserve_previous_output(self) -> None:
        corrections = (
            NameCorrection("CARLOS EJEMPLO", "Carlos Prueba", "Referencia A"),
            NameCorrection(" Carlos  Ejemplo ", "Carlos Otro", "Referencia B"),
        )
        previous_content = TEMPLATE_DOCUMENT.read_bytes()
        self.output_path.write_bytes(previous_content)

        with patch("meeting_generator.generator.fill_template") as writer:
            result = generate_document(
                SOURCE_DOCUMENT,
                TEMPLATE_DOCUMENT,
                self.output_path,
                name_corrections=corrections,
            )

        writer.assert_not_called()
        self.assertFalse(result.succeeded)
        self.assertEqual(result.written_weeks, 0)
        self.assertEqual(result.omitted_weeks, len(self.valid_weeks))
        self.assertTrue(any("repiten" in error for error in result.errors))
        self.assertEqual(self.output_path.read_bytes(), previous_content)

    def test_one_table_error_prevents_all_other_weeks_from_being_written(
        self,
    ) -> None:
        original_parse_week = ProgramParser._parse_week_from_table

        def fail_second_table(parser, table, week_number, header_data):
            if week_number == 2:
                raise ValueError("tabla dañada para la prueba")
            return original_parse_week(parser, table, week_number, header_data)

        with patch.object(
            ProgramParser,
            "_parse_week_from_table",
            new=fail_second_table,
        ):
            result = generate_document(
                SOURCE_DOCUMENT,
                TEMPLATE_DOCUMENT,
                self.output_path,
            )

        self.assertFalse(result.succeeded)
        self.assertEqual(result.written_weeks, 0)
        self.assertEqual(result.omitted_weeks, len(self.valid_weeks))
        self.assertTrue(
            any("tabla 2" in error.lower() for error in result.errors),
            result.errors,
        )
        self.assertIsNone(result.output_path)
        self.assertFalse(self.output_path.exists())

    def test_failed_save_does_not_claim_or_replace_preexisting_output(self) -> None:
        previous_content = b"archivo anterior"
        self.output_path.write_bytes(previous_content)

        with patch("docx.document.Document.save", autospec=True):
            result = fill_template(
                TEMPLATE_DOCUMENT,
                [self.valid_weeks[0]],
                self.output_path,
            )

        self.assertFalse(result.succeeded)
        self.assertEqual(result.written_weeks, 0)
        self.assertEqual(result.omitted_weeks, 1)
        self.assertTrue(result.errors)
        self.assertIsNone(result.output_path)
        self.assertEqual(self.output_path.read_bytes(), previous_content)

    def test_ambiguous_combined_row_never_creates_or_replaces_output(self) -> None:
        directory = Path(self.temp_directory.name)
        source_path = directory / "ambiguous-source.docx"
        previous_content = TEMPLATE_DOCUMENT.read_bytes()
        for column in (1, 2):
            for existing_output in (False, True):
                with self.subTest(column=column, existing_output=existing_output):
                    document = Document(SOURCE_DOCUMENT)
                    document.tables[7].rows[10].cells[column].add_paragraph(
                        "Dato adicional"
                    )
                    document.save(source_path)
                    output_path = directory / f"output-{column}-{existing_output}.docx"
                    if existing_output:
                        output_path.write_bytes(previous_content)
                    files_before = set(directory.iterdir())

                    with patch("meeting_generator.generator.fill_template") as writer:
                        result = generate_document(
                            source_path,
                            TEMPLATE_DOCUMENT,
                            output_path,
                        )

                    writer.assert_not_called()
                    self.assertFalse(result.succeeded)
                    self.assertEqual(result.written_weeks, 0)
                    self.assertEqual(result.omitted_weeks, len(self.valid_weeks))
                    self.assertIsNone(result.output_path)
                    self.assertTrue(
                        any("semana 8, fila 11" in error for error in result.errors),
                        result.errors,
                    )
                    self.assertEqual(set(directory.iterdir()), files_before)
                    if existing_output:
                        self.assertEqual(output_path.read_bytes(), previous_content)
                    else:
                        self.assertFalse(output_path.exists())

    def test_missing_assignment_number_never_creates_or_replaces_output(self) -> None:
        directory = Path(self.temp_directory.name)
        source_path = directory / "missing-number-source.docx"
        document = Document(SOURCE_DOCUMENT)
        table = document.tables[0]
        row_index, row = next(
            (index, row)
            for index, row in enumerate(table.rows, start=1)
            if row.cells[0].text == "4"
        )
        row.cells[0].text = ""
        document.save(source_path)
        original_source = source_path.read_bytes()
        original_template = TEMPLATE_DOCUMENT.read_bytes()

        for existing_output in (False, True):
            with self.subTest(existing_output=existing_output):
                output_path = directory / f"output-{existing_output}.docx"
                if existing_output:
                    output_path.write_bytes(original_template)
                files_before = set(directory.iterdir())

                with patch("meeting_generator.generator.fill_template") as writer:
                    result = generate_document(
                        source_path, TEMPLATE_DOCUMENT, output_path
                    )

                writer.assert_not_called()
                self.assertFalse(result.succeeded)
                self.assertEqual(result.written_weeks, 0)
                self.assertEqual(result.omitted_weeks, len(self.valid_weeks))
                self.assertIsNone(result.output_path)
                self.assertTrue(
                    any(
                        f"semana 1, fila {row_index}" in error
                        and "falta el número" in error
                        for error in result.errors
                    ),
                    result.errors,
                )
                self.assertEqual(set(directory.iterdir()), files_before)
                self.assertEqual(source_path.read_bytes(), original_source)
                self.assertEqual(TEMPLATE_DOCUMENT.read_bytes(), original_template)
                if existing_output:
                    self.assertEqual(output_path.read_bytes(), original_template)
                else:
                    self.assertFalse(output_path.exists())

    def test_output_collisions_never_modify_source_or_template(self) -> None:
        original_source = SOURCE_DOCUMENT.read_bytes()
        original_template = TEMPLATE_DOCUMENT.read_bytes()

        for protected_path, expected_error in (
            (SOURCE_DOCUMENT, "fuente"),
            (TEMPLATE_DOCUMENT, "plantilla"),
        ):
            with self.subTest(output=protected_path.name):
                result = generate_document(
                    SOURCE_DOCUMENT,
                    TEMPLATE_DOCUMENT,
                    protected_path,
                )

                self.assertFalse(result.succeeded)
                self.assertEqual(result.written_weeks, 0)
                self.assertGreater(result.omitted_weeks, 0)
                self.assertIsNone(result.output_path)
                self.assertTrue(
                    any(expected_error in error.lower() for error in result.errors),
                    result.errors,
                )

        self.assertEqual(SOURCE_DOCUMENT.read_bytes(), original_source)
        self.assertEqual(TEMPLATE_DOCUMENT.read_bytes(), original_template)

    def test_success_is_published_with_atomic_replace_in_output_directory(
        self,
    ) -> None:
        with patch(
            "meeting_generator.template_writer.os.replace",
            wraps=os.replace,
        ) as replace:
            result = fill_template(
                TEMPLATE_DOCUMENT,
                [self.valid_weeks[0]],
                self.output_path,
            )

        self.assertTrue(result.succeeded)
        replace.assert_called_once()
        temporary_path, published_path = replace.call_args.args
        self.assertEqual(Path(published_path), self.output_path)
        self.assertEqual(Path(temporary_path).parent, self.output_path.parent)
        self.assertNotEqual(Path(temporary_path), self.output_path)
        self.assertFalse(Path(temporary_path).exists())

    def test_failed_atomic_replace_preserves_output_and_removes_temporary(self) -> None:
        previous_content = b"archivo anterior"
        self.output_path.write_bytes(previous_content)
        files_before = set(self.output_path.parent.iterdir())

        with patch(
            "meeting_generator.template_writer.os.replace",
            side_effect=OSError("reemplazo interrumpido"),
        ):
            result = fill_template(
                TEMPLATE_DOCUMENT,
                [self.valid_weeks[0]],
                self.output_path,
            )

        self.assertFalse(result.succeeded)
        self.assertIsNone(result.output_path)
        self.assertEqual(self.output_path.read_bytes(), previous_content)
        self.assertEqual(set(self.output_path.parent.iterdir()), files_before)


class GuiGenerationResultTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_directory = tempfile.TemporaryDirectory()
        directory = Path(self.temp_directory.name)
        self.output_path = directory / "S-140_COMPLETADO.docx"

        self.app = MeetingGeneratorApp.__new__(MeetingGeneratorApp)
        self.app.source_path = directory / "source.docx"
        self.app.template_path = directory / "template.docx"
        self.app._generation_in_progress = False
        self.app._executor = ThreadPoolExecutor(max_workers=1)
        self.app._generation_future = None
        self.app._generation_poll_id = None
        self.app._close_requested = False
        self.app.root = Mock()
        self.app.status_label = Mock()
        self.app.source_label = Mock()
        self.app.template_label = Mock()
        self.app.source_button = Mock()
        self.app.template_button = Mock()
        self.app.generate_button = Mock()
        self.app.progress = Mock()
        self.app.error_panel = Mock()
        self.app.settings = AppSettings(directory / "settings.json")
        self.ui_thread = get_ident()
        self.after_callbacks: dict[str, Callable[[], None]] = {}
        self.next_after_id = 0
        self.app.root.after.side_effect = self._schedule_after
        self.app.root.after_cancel.side_effect = self._cancel_after
        for method in (
            self.app.root.update_idletasks,
            self.app.root.destroy,
            self.app.status_label.config,
            self.app.source_label.config,
            self.app.template_label.config,
            self.app.source_button.config,
            self.app.template_button.config,
            self.app.generate_button.config,
        ):
            method.side_effect = self._assert_ui_thread

    def tearDown(self) -> None:
        self.app._executor.shutdown(wait=True, cancel_futures=True)
        self.temp_directory.cleanup()

    def _assert_ui_thread(self, *_args, **_kwargs) -> None:
        self.assertEqual(get_ident(), self.ui_thread)

    def _schedule_after(self, _delay: int, callback, *args) -> str:
        self._assert_ui_thread()
        self.next_after_id += 1
        callback_id = f"after-{self.next_after_id}"
        self.after_callbacks[callback_id] = lambda: callback(*args)
        return callback_id

    def _cancel_after(self, callback_id: str) -> None:
        self._assert_ui_thread()
        self.after_callbacks.pop(callback_id)

    def _finish_pending_generation(self) -> None:
        future = self.app._generation_future
        if future is not None:
            completed, _pending = wait((future,), timeout=10)
            self.assertIn(future, completed, "El trabajador no terminó")
            callback = self.after_callbacks.pop(self.app._generation_poll_id)
            callback()
            self.assertFalse(self.app._generation_in_progress)
            self.assertIsNone(self.app._generation_future)
            self.assertIsNone(self.app._generation_poll_id)
            self.assertFalse(self.after_callbacks)

    def _generate_and_wait(self) -> None:
        self.app._generate_document()
        self._finish_pending_generation()

    def test_canceling_file_selection_preserves_paths_labels_and_button_state(
        self,
    ) -> None:
        source_path = self.app.source_path
        template_path = self.app.template_path
        for source_selected, template_selected in (
            (False, False),
            (True, False),
            (False, True),
            (True, True),
        ):
            for selector in (
                self.app._select_source_file,
                self.app._select_template_file,
            ):
                with self.subTest(
                    source=source_selected,
                    template=template_selected,
                    selector=selector.__name__,
                ):
                    self.app.source_path = source_path if source_selected else None
                    self.app.template_path = (
                        template_path if template_selected else None
                    )
                    source_label = {
                        "text": f"✓ {source_path.name}"
                        if source_selected
                        else "No se ha seleccionado archivo",
                        "fg": COLOR_SUCCESS if source_selected else COLOR_MUTED,
                    }
                    template_label = {
                        "text": f"✓ {template_path.name}"
                        if template_selected
                        else "No se ha seleccionado archivo",
                        "fg": COLOR_SUCCESS if template_selected else COLOR_MUTED,
                    }
                    self.app.source_label.config.side_effect = source_label.update
                    self.app.template_label.config.side_effect = template_label.update
                    previous_labels = (source_label.copy(), template_label.copy())
                    previous_paths = (self.app.source_path, self.app.template_path)
                    expected_state = (
                        tk.NORMAL
                        if source_selected and template_selected
                        else tk.DISABLED
                    )
                    self.app._update_generate_button()

                    with patch(
                        "meeting_generator.main.filedialog.askopenfilename",
                        return_value="",
                    ):
                        selector()

                    self.assertEqual(
                        (self.app.source_path, self.app.template_path), previous_paths
                    )
                    self.assertEqual((source_label, template_label), previous_labels)
                    self.assertEqual(
                        self.app.generate_button.config.call_args.kwargs["state"],
                        expected_state,
                    )

    def test_generation_uses_latest_visible_selection_after_canceled_dialogs(
        self,
    ) -> None:
        self.app.source_path = None
        self.app.template_path = None
        source_label = {"text": "No se ha seleccionado archivo", "fg": COLOR_MUTED}
        template_label = source_label.copy()
        self.app.source_label.config.side_effect = source_label.update
        self.app.template_label.config.side_effect = template_label.update

        with patch(
            "meeting_generator.main.filedialog.askopenfilename",
            side_effect=(
                str(SOURCE_DOCUMENT),
                str(TEMPLATE_DOCUMENT),
                "",
                str(SECOND_SOURCE_DOCUMENT),
                "",
            ),
        ):
            self.app._select_source_file()
            self.assertEqual(
                self.app.generate_button.config.call_args.kwargs["state"], tk.DISABLED
            )
            self.app._select_template_file()
            self.app._select_source_file()
            self.assertEqual(source_label["text"], f"✓ {SOURCE_DOCUMENT.name}")
            self.app._select_source_file()
            self.app._select_template_file()

        self.assertEqual(
            source_label, {"text": f"✓ {SECOND_SOURCE_DOCUMENT.name}", "fg": COLOR_SUCCESS}
        )
        self.assertEqual(
            template_label, {"text": f"✓ {TEMPLATE_DOCUMENT.name}", "fg": COLOR_SUCCESS}
        )
        self.assertEqual(
            self.app.generate_button.config.call_args.kwargs["state"], tk.NORMAL
        )
        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ),
            patch(
                "meeting_generator.main.generate_document", wraps=generate_document
            ) as generate,
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        generate.assert_called_once_with(
            SECOND_SOURCE_DOCUMENT, TEMPLATE_DOCUMENT, self.output_path
        )
        show_error.assert_not_called()
        show_info.assert_called_once()
        self.assertTrue(self.output_path.is_file())

    def test_gui_success_uses_result_path_and_written_count(self) -> None:
        self.output_path.write_bytes(b"documento generado")
        result = GenerationResult.success(3, self.output_path)

        with (
            patch(
                "meeting_generator.main.get_default_output_path",
                return_value=self.output_path,
            ),
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ),
            patch("meeting_generator.main.messagebox.askyesno", return_value=True),
            patch("meeting_generator.main.generate_document", return_value=result),
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        show_error.assert_not_called()
        show_info.assert_called_once()
        success_message = show_info.call_args.args[1]
        self.assertIn(self.output_path.name, success_message)
        self.assertIn("3 semanas", success_message)
        self.app.root.update.assert_not_called()

    def test_gui_never_announces_failure_as_success_for_preexisting_path(
        self,
    ) -> None:
        self.output_path.write_bytes(b"archivo anterior")
        result = GenerationResult.failure(
            4,
            ["La plantilla no tiene capacidad suficiente"],
        )

        with (
            patch(
                "meeting_generator.main.get_default_output_path",
                return_value=self.output_path,
            ),
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ),
            patch("meeting_generator.main.messagebox.askyesno", return_value=True),
            patch("meeting_generator.main.generate_document", return_value=result),
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        show_info.assert_not_called()
        show_error.assert_called_once()
        failure_message = show_error.call_args.args[1]
        self.assertIn("4", failure_message)
        self.assertIn("capacidad", failure_message)

    def test_gui_reports_publication_failure_without_accessing_the_output(
        self,
    ) -> None:
        for existing_output in (False, True):
            for available_output in (False, True):
                with self.subTest(
                    existing_output=existing_output,
                    available_output=available_output,
                ):
                    if available_output:
                        self.output_path.write_bytes(b"archivo publicado")
                    else:
                        self.output_path.unlink(missing_ok=True)
                    result = GenerationResult.published_failure(
                        3,
                        self.output_path,
                        ["Error al leer el archivo publicado"],
                    )

                    with (
                        patch.object(
                            Path, "is_file", side_effect=PermissionError("sin acceso")
                        ) as is_file,
                        patch.object(
                            Path, "stat", side_effect=PermissionError("sin acceso")
                        ) as stat,
                        patch(
                            "meeting_generator.main.messagebox.showinfo"
                        ) as show_info,
                        patch(
                            "meeting_generator.main.messagebox.showerror",
                            side_effect=self._assert_ui_thread,
                        ) as show_error,
                    ):
                        self.app._show_generation_result(result, existing_output)

                    is_file.assert_not_called()
                    stat.assert_not_called()
                    show_info.assert_not_called()
                    show_error.assert_called_once()
                    failure_message = show_error.call_args.args[1]
                    self.assertIn("El archivo se publicó", failure_message)
                    self.assertIn(str(self.output_path), failure_message)
                    self.assertIn("posterior a la publicación", failure_message)
                    self.assertIn("Semanas escritas: 3", failure_message)
                    self.assertIn("Semanas omitidas: 0", failure_message)
                    self.assertIn("Error al leer", failure_message)
                    self.assertIn(
                        "El archivo anterior fue reemplazado"
                        if existing_output
                        else "El archivo fue creado",
                        failure_message,
                    )
                    self.assertNotIn("no corresponde", failure_message)
                    self.assertNotIn("No se confirmó un archivo nuevo", failure_message)
                    status = self.app.status_label.config.call_args.kwargs
                    self.assertEqual(status["fg"], COLOR_ERROR)
                    self.assertIn("Archivo publicado con errores", status["text"])
                    self.assertIn("Semanas escritas: 3", status["text"])
                    self.assertIn("Semanas omitidas: 0", status["text"])

    def test_gui_reports_publication_even_if_logging_the_error_fails(self) -> None:
        result = GenerationResult.published_failure(
            3,
            self.output_path,
            ["No se pudo verificar el archivo publicado"],
        )
        with (
            patch(
                "meeting_generator.main.logger.error",
                side_effect=OSError("log no disponible"),
            ) as log_error,
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self.app._show_generation_result(result, output_existed_before=True)

        log_error.assert_called_once()
        show_info.assert_not_called()
        show_error.assert_called_once()
        failure_message = show_error.call_args.args[1]
        self.assertIn("El archivo se publicó", failure_message)
        self.assertIn(str(self.output_path), failure_message)
        self.assertIn("El archivo anterior fue reemplazado", failure_message)
        self.assertIn("Semanas escritas: 3", failure_message)
        self.assertIn("Semanas omitidas: 0", failure_message)
        self.assertIn(
            "No se pudo registrar el error en el log: log no disponible",
            failure_message,
        )
        self.assertEqual(self.app.status_label.config.call_args.kwargs["fg"], COLOR_ERROR)

    def test_gui_reports_published_output_that_can_no_longer_be_verified(self) -> None:
        for existing_output in (False, True):
            for unavailable_state in ("deleted", "empty", "inaccessible"):
                with self.subTest(
                    existing_output=existing_output,
                    unavailable_state=unavailable_state,
                ):
                    self.output_path.write_bytes(b"archivo publicado")
                    result = GenerationResult.success(3, self.output_path)
                    if unavailable_state == "deleted":
                        self.output_path.unlink()
                    elif unavailable_state == "empty":
                        self.output_path.write_bytes(b"")
                    original_stat = Path.stat

                    def stat(path, *args, **kwargs):
                        if (
                            path == self.output_path
                            and unavailable_state == "inaccessible"
                        ):
                            raise PermissionError("sin acceso al archivo publicado")
                        return original_stat(path, *args, **kwargs)

                    with (
                        patch.object(Path, "stat", new=stat),
                        patch(
                            "meeting_generator.main.messagebox.showinfo"
                        ) as show_info,
                        patch(
                            "meeting_generator.main.messagebox.showerror"
                        ) as show_error,
                    ):
                        self.app._show_generation_result(result, existing_output)

                    show_info.assert_not_called()
                    show_error.assert_called_once()
                    failure_message = show_error.call_args.args[1]
                    self.assertIn("El archivo se publicó", failure_message)
                    self.assertIn(
                        "No se pudo verificar el archivo publicado", failure_message
                    )
                    self.assertIn("Semanas escritas: 3", failure_message)
                    self.assertIn("Semanas omitidas: 0", failure_message)
                    self.assertIn(
                        "El archivo anterior fue reemplazado"
                        if existing_output
                        else "El archivo fue creado",
                        failure_message,
                    )
                    self.assertNotIn("no corresponde", failure_message)
                    self.assertEqual(
                        self.app.status_label.config.call_args.kwargs["fg"], COLOR_ERROR
                    )

    def test_gui_reports_read_error_after_real_worker_publishes_the_document(
        self,
    ) -> None:
        self.app.source_path = SOURCE_DOCUMENT
        self.app.template_path = TEMPLATE_DOCUMENT
        original_digest = TemplateWriter._file_digest
        weeks = len(parse_document(SOURCE_DOCUMENT))

        def fail_published_digest(writer, path):
            if path == self.output_path:
                raise OSError("lectura del archivo publicado interrumpida")
            return original_digest(writer, path)

        for existing_output in (False, True):
            with self.subTest(existing_output=existing_output):
                if existing_output:
                    self.output_path.write_bytes(b"archivo anterior")
                else:
                    self.output_path.unlink(missing_ok=True)
                with (
                    patch(
                        "meeting_generator.main.filedialog.asksaveasfilename",
                        return_value=str(self.output_path),
                    ),
                    patch(
                        "meeting_generator.main.messagebox.askyesno", return_value=True
                    ),
                    patch.object(
                        TemplateWriter,
                        "_file_digest",
                        autospec=True,
                        side_effect=fail_published_digest,
                    ),
                    patch("meeting_generator.main.messagebox.showinfo") as show_info,
                    patch(
                        "meeting_generator.main.messagebox.showerror",
                        side_effect=self._assert_ui_thread,
                    ) as show_error,
                ):
                    self._generate_and_wait()

                show_info.assert_not_called()
                show_error.assert_called_once()
                failure_message = show_error.call_args.args[1]
                self.assertIn("El archivo se publicó", failure_message)
                self.assertIn(str(self.output_path), failure_message)
                self.assertIn(
                    "lectura del archivo publicado interrumpida", failure_message
                )
                self.assertIn(f"Semanas escritas: {weeks}", failure_message)
                self.assertIn("Semanas omitidas: 0", failure_message)
                self.assertIn(
                    "El archivo anterior fue reemplazado"
                    if existing_output
                    else "El archivo fue creado",
                    failure_message,
                )
                self.assertNotIn("no corresponde", failure_message)
                self.assertTrue(Document(self.output_path).tables)
                self.assertEqual(
                    self.app.status_label.config.call_args.kwargs["fg"], COLOR_ERROR
                )

    def test_gui_shows_combined_row_location_and_preserves_previous_output(
        self,
    ) -> None:
        document = Document(SOURCE_DOCUMENT)
        document.tables[7].rows[10].cells[2].add_paragraph("Ana Ejemplo")
        document.save(self.app.source_path)
        self.app.template_path = TEMPLATE_DOCUMENT
        previous_content = TEMPLATE_DOCUMENT.read_bytes()
        self.output_path.write_bytes(previous_content)

        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ),
            patch("meeting_generator.main.messagebox.askyesno", return_value=True),
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        show_info.assert_not_called()
        show_error.assert_called_once()
        failure_message = show_error.call_args.args[1]
        self.assertIn("semana 8, fila 11", failure_message)
        self.assertIn(
            "números=2, títulos=2, grupos de participantes=3", failure_message
        )
        self.assertEqual(self.output_path.read_bytes(), previous_content)
        self.assertFalse(self.app._generation_in_progress)

    def test_canceling_save_as_does_not_start_generation(self) -> None:
        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value="",
            ) as save_as,
            patch("meeting_generator.main.generate_document") as generate,
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        save_as.assert_called_once()
        save_options = save_as.call_args.kwargs
        self.assertEqual(save_options["defaultextension"], ".docx")
        self.assertEqual(save_options["initialfile"], "S-140_COMPLETADO.docx")
        generate.assert_not_called()
        show_info.assert_not_called()
        show_error.assert_not_called()
        self.assertFalse(self.app._generation_in_progress)
        self.app.root.update.assert_not_called()

    def test_generation_uses_paths_captured_before_save_dialog(self) -> None:
        source_snapshot = self.app.source_path
        template_snapshot = self.app.template_path
        changed_source = source_snapshot.with_name("changed-source.docx")
        changed_template = template_snapshot.with_name("changed-template.docx")

        def change_gui_state_during_dialog(**_kwargs) -> str:
            self.app.source_path = changed_source
            self.app.template_path = changed_template
            return str(self.output_path)

        def generate_from_snapshots(
            source_path: Path,
            template_path: Path,
            output_path: Path,
        ) -> GenerationResult:
            self.output_path.write_bytes(b"documento generado")
            return GenerationResult.success(1, output_path)

        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                side_effect=change_gui_state_during_dialog,
            ),
            patch(
                "meeting_generator.main.generate_document",
                side_effect=generate_from_snapshots,
            ) as generate,
            patch("meeting_generator.main.messagebox.showinfo"),
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        show_error.assert_not_called()
        generate.assert_called_once_with(
            source_snapshot,
            template_snapshot,
            self.output_path,
        )

    def test_reentrant_generation_is_ignored(self) -> None:
        self.app._generation_in_progress = True

        with (
            patch("meeting_generator.main.filedialog.asksaveasfilename") as save_as,
            patch("meeting_generator.main.generate_document") as generate,
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        save_as.assert_not_called()
        generate.assert_not_called()
        show_info.assert_not_called()
        show_error.assert_not_called()
        self.assertTrue(self.app._generation_in_progress)
        self.app.root.update.assert_not_called()

    def test_controls_are_disabled_during_work_and_restored_afterwards(self) -> None:
        def states(control: Mock) -> list[str]:
            return [
                call.kwargs["state"]
                for call in control.config.call_args_list
                if "state" in call.kwargs
            ]

        def generate_while_disabled(
            _source_path: Path,
            _template_path: Path,
            output_path: Path,
        ) -> GenerationResult:
            self.assertTrue(self.app._generation_in_progress)
            for control in (
                self.app.source_button,
                self.app.template_button,
                self.app.generate_button,
            ):
                self.assertEqual(states(control)[-1], tk.DISABLED)
            self.output_path.write_bytes(b"documento generado")
            return GenerationResult.success(1, output_path)

        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ),
            patch(
                "meeting_generator.main.generate_document",
                side_effect=generate_while_disabled,
            ),
            patch("meeting_generator.main.messagebox.showinfo"),
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        show_error.assert_not_called()
        self.assertFalse(self.app._generation_in_progress)
        for control in (
            self.app.source_button,
            self.app.template_button,
            self.app.generate_button,
        ):
            control_states = states(control)
            self.assertIn(tk.DISABLED, control_states)
            self.assertEqual(control_states[-1], tk.NORMAL)
        self.app.root.update.assert_not_called()

    def test_existing_output_requires_confirmation_before_overwrite(self) -> None:
        previous_content = b"archivo anterior"
        self.output_path.write_bytes(previous_content)

        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ),
            patch(
                "meeting_generator.main.messagebox.askyesno",
                return_value=False,
            ) as confirm,
            patch("meeting_generator.main.generate_document") as generate,
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        confirm.assert_called_once()
        generate.assert_not_called()
        show_info.assert_not_called()
        show_error.assert_not_called()
        self.assertEqual(self.output_path.read_bytes(), previous_content)
        self.assertFalse(self.app._generation_in_progress)

    def test_gui_rejects_output_collisions_before_generation(self) -> None:
        for protected_path, expected_text in (
            (self.app.source_path, "fuente"),
            (self.app.template_path, "plantilla"),
        ):
            with (
                self.subTest(output=protected_path),
                patch(
                    "meeting_generator.main.filedialog.asksaveasfilename",
                    return_value=str(protected_path),
                ),
                patch("meeting_generator.main.messagebox.askyesno") as confirm,
                patch("meeting_generator.main.generate_document") as generate,
                patch("meeting_generator.main.messagebox.showerror") as show_error,
            ):
                self._generate_and_wait()

            confirm.assert_not_called()
            generate.assert_not_called()
            show_error.assert_called_once()
            self.assertIn(expected_text, show_error.call_args.args[1])
            self.assertFalse(self.app._generation_in_progress)

    def test_controls_are_restored_when_generation_raises(self) -> None:
        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ),
            patch(
                "meeting_generator.main.generate_document",
                side_effect=RuntimeError("fallo inesperado"),
            ),
            patch(
                "meeting_generator.main.messagebox.showinfo",
                side_effect=self._assert_ui_thread,
            ) as show_info,
            patch(
                "meeting_generator.main.messagebox.showerror",
                side_effect=self._assert_ui_thread,
            ) as show_error,
        ):
            self._generate_and_wait()

        show_info.assert_not_called()
        show_error.assert_called_once()
        self.assertIn("fallo inesperado", show_error.call_args.args[1])
        self.assertEqual(self.app.status_label.config.call_args.kwargs["fg"], COLOR_ERROR)
        self.assertFalse(self.app._generation_in_progress)
        for control in (
            self.app.source_button,
            self.app.template_button,
            self.app.generate_button,
        ):
            restored_states = [
                call.kwargs["state"]
                for call in control.config.call_args_list
                if "state" in call.kwargs
            ]
            self.assertEqual(restored_states[-1], tk.NORMAL)

    def test_generation_keeps_event_loop_responsive_and_uses_one_worker(self) -> None:
        started = Event()
        release = Event()
        worker_threads: list[int] = []

        def generate_in_worker(
            _source_path: Path, _template_path: Path, output_path: Path
        ) -> GenerationResult:
            worker_threads.append(get_ident())
            started.set()
            if not release.wait(timeout=10):
                raise RuntimeError("El trabajador no fue liberado por la prueba")
            output_path.write_bytes(b"documento generado")
            return GenerationResult.success(1, output_path)

        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ) as save_as,
            patch("meeting_generator.main.messagebox.askyesno", return_value=True),
            patch("meeting_generator.main.filedialog.askopenfilename") as open_file,
            patch(
                "meeting_generator.main.generate_document",
                side_effect=generate_in_worker,
            ) as generate,
            patch(
                "meeting_generator.main.messagebox.showinfo",
                side_effect=self._assert_ui_thread,
            ) as show_info,
            patch(
                "meeting_generator.main.messagebox.showerror",
                side_effect=self._assert_ui_thread,
            ) as show_error,
        ):
            try:
                self.app._generate_document()
                self.assertTrue(started.wait(timeout=5))
                self.assertNotEqual(worker_threads[0], self.ui_thread)
                self.assertTrue(self.app._generation_in_progress)
                original_future = self.app._generation_future

                # Un evento de consulta vuelve al bucle aunque siga trabajando.
                callback = self.after_callbacks.pop(self.app._generation_poll_id)
                callback()
                self.assertTrue(self.after_callbacks)
                probe = Mock()
                probe_id = self.app.root.after(0, probe)
                self.after_callbacks.pop(probe_id)()
                probe.assert_called_once()

                self.app._generate_document()
                self.app._select_source_file()
                self.app._select_template_file()
                self.assertIs(self.app._generation_future, original_future)
                save_as.assert_called_once()
                generate.assert_called_once()
                open_file.assert_not_called()
                show_info.assert_not_called()
                show_error.assert_not_called()
                for control in (
                    self.app.source_button,
                    self.app.template_button,
                    self.app.generate_button,
                ):
                    self.assertEqual(
                        control.config.call_args.kwargs["state"], tk.DISABLED
                    )
            finally:
                release.set()

            self._finish_pending_generation()
            self.assertFalse(self.app._generation_in_progress)
            self.assertFalse(self.after_callbacks)
            show_info.assert_called_once()
            self._generate_and_wait()

        self.assertEqual(len(worker_threads), 2)
        self.assertEqual(worker_threads[0], worker_threads[1])
        self.assertEqual(show_info.call_count, 2)
        show_error.assert_not_called()

    def test_controls_are_restored_when_worker_cannot_start(self) -> None:
        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ),
            patch.object(
                self.app._executor, "submit", side_effect=RuntimeError("sin trabajador")
            ),
            patch("meeting_generator.main.generate_document") as generate,
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch(
                "meeting_generator.main.messagebox.showerror",
                side_effect=self._assert_ui_thread,
            ) as show_error,
        ):
            self._generate_and_wait()

        generate.assert_not_called()
        show_info.assert_not_called()
        show_error.assert_called_once()
        self.assertIn("sin trabajador", show_error.call_args.args[1])
        self.app.root.after_cancel.assert_called_once()
        self.assertFalse(self.after_callbacks)
        self.assertIsNone(self.app._generation_poll_id)
        self.assertIsNone(self.app._generation_future)
        self.assertFalse(self.app._generation_in_progress)
        for control in (
            self.app.source_button,
            self.app.template_button,
            self.app.generate_button,
        ):
            self.assertEqual(control.config.call_args.kwargs["state"], tk.NORMAL)

    def test_closing_waits_for_generation_without_completion_dialogs(self) -> None:
        started = Event()
        release = Event()

        def generate_before_closing(
            _source_path: Path, _template_path: Path, output_path: Path
        ) -> GenerationResult:
            started.set()
            if not release.wait(timeout=10):
                raise RuntimeError("El trabajador no fue liberado por la prueba")
            output_path.write_bytes(b"documento generado antes del cierre")
            return GenerationResult.success(1, output_path)

        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                return_value=str(self.output_path),
            ) as save_as,
            patch(
                "meeting_generator.main.generate_document",
                side_effect=generate_before_closing,
            ),
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            try:
                self.app._generate_document()
                self.assertTrue(started.wait(timeout=5))
                self.app._close()
                self.app.root.destroy.assert_not_called()
                self.assertTrue(self.app._generation_in_progress)
                self.assertIn(
                    "Cerrando", self.app.status_label.config.call_args.kwargs["text"]
                )
                self.app._generate_document()
                save_as.assert_called_once()
            finally:
                release.set()
            self._finish_pending_generation()

        show_info.assert_not_called()
        show_error.assert_not_called()
        self.assertEqual(
            self.output_path.read_bytes(), b"documento generado antes del cierre"
        )
        self.app.root.destroy.assert_called_once()
        self.assertFalse(self.app._generation_in_progress)
        self.assertFalse(self.after_callbacks)
        self.assertIsNone(self.app._generation_future)
        with self.assertRaises(RuntimeError):
            self.app._executor.submit(lambda: None)

    def test_closing_during_save_dialog_does_not_start_worker(self) -> None:
        def close_during_dialog(**_kwargs) -> str:
            self.app._close()
            return str(self.output_path)

        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename",
                side_effect=close_during_dialog,
            ),
            patch("meeting_generator.main.generate_document") as generate,
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self._generate_and_wait()

        generate.assert_not_called()
        show_info.assert_not_called()
        show_error.assert_not_called()
        self.app.root.destroy.assert_called_once()
        self.assertFalse(self.app._generation_in_progress)
        self.assertFalse(self.after_callbacks)


class NameNormalizationTests(unittest.TestCase):
    def test_corrections_require_original_replacement_and_reason(self) -> None:
        for field in ("original_name", "corrected_name", "reason"):
            for invalid_value in ("", " \t\n ", None):
                with self.subTest(field=field, value=invalid_value):
                    values = {
                        "original_name": "Ana Ejemplo",
                        "corrected_name": "Ana Prueba",
                        "reason": "Verificado en la referencia sintética",
                    }
                    values[field] = invalid_value
                    with self.assertRaises(ValueError):
                        NameCorrection(**values)

    def test_explicit_corrections_match_full_names_and_are_logged(self) -> None:
        corrections = (
            NameCorrection(
                "MIRZA CERRANO(h)",
                "Mirza Serrano (h)",
                "Escritura confirmada en la referencia sintética",
            ),
            NameCorrection(
                "Sussy Ejemplo", "Susy Ejemplo", "Nombre confirmado en la referencia"
            ),
            NameCorrection(
                "Beatriz Ejemplo",
                "Berta Ejemplo",
                "Ayudante confirmada en la referencia",
            ),
        )
        with tempfile.TemporaryDirectory() as directory:
            source_path = Path(directory) / "source.docx"
            document = Document(SOURCE_DOCUMENT)
            for paragraph in document.paragraphs:
                if "PRESIDENTE:" in paragraph.text:
                    paragraph.text = paragraph.text.replace(
                        "ANA EJEMPLO", "SUSSY EJEMPLO"
                    )
            document.tables[0].rows[1].cells[2].text = "MIRZA CERRANO(h)"
            document.tables[0].rows[2].cells[2].text = "MIRZA CERRANO LÓPEZ (h)"
            document.tables[0].rows[-1].cells[2].text = "SUSSY EJEMPLO"
            document.save(source_path)
            original_source = source_path.read_bytes()

            with self.assertLogs("meeting_generator.parser", level="INFO") as logs:
                weeks = parse_document(source_path, name_corrections=corrections)

            self.assertEqual(source_path.read_bytes(), original_source)
            correction_logs = [
                record.getMessage()
                for record in logs.records
                if "Corrección explícita de nombre" in record.getMessage()
            ]
            self.assertTrue(
                any(
                    f"fuente={source_path}" in message
                    and "semana=1, campo=punto 1, participante 1" in message
                    and "original='Mirza Cerrano (h)'" in message
                    and "corregido='Mirza Serrano (h)'" in message
                    and "motivo=Escritura confirmada en la referencia sintética"
                    in message
                    for message in correction_logs
                ),
                correction_logs,
            )

        first_week = weeks[0]
        self.assertEqual(first_week.president, "Susy Ejemplo")
        self.assertEqual(first_week.opening_prayer, "Susy Ejemplo")
        self.assertEqual(first_week.closing_prayer, "Susy Ejemplo")
        self.assertEqual(
            first_week.bible_treasures[0].first_participant_name(), "Mirza Serrano (h)"
        )
        self.assertEqual(
            first_week.bible_treasures[1].first_participant_name(),
            "Mirza Cerrano López (h)",
        )
        self.assertEqual(
            first_week.ministry[0].first_participant_name(), "Alicia Ejemplo"
        )
        self.assertEqual(
            first_week.ministry[0].second_participant_name(), "Berta Ejemplo"
        )

    def test_normalization_preserves_spelling_punctuation_and_full_names(self) -> None:
        cases = (
            ("  MIRZA\tCERRANO (h)  ", "Mirza Cerrano (h)"),
            ("SUSSY BANEGAS", "Sussy Banegas"),
            ("JAMES O”CONNOR", "James O”Connor"),
            ("Susana Cerrano", "Susana Cerrano"),
            ("Sussy Ejemplo", "Sussy Ejemplo"),
            ("JORGE DANIEL VALLADARES(p)", "Jorge Daniel Valladares (p)"),
        )
        for original, expected in cases:
            with self.subTest(name=original):
                self.assertEqual(normalize_person_name(original), expected)

        self.assertEqual(
            split_names("MIRZA CERRANO (h) // SUSSY BANEGAS (h)"),
            ("Mirza Cerrano (h)", "Sussy Banegas (h)"),
        )

    def test_parser_preserves_names_without_explicit_corrections(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source_path = Path(directory) / "source.docx"
            document = Document(SECOND_SOURCE_DOCUMENT)
            for paragraph in document.paragraphs:
                if "JEREMÍAS 45-46" in paragraph.text:
                    paragraph.text = paragraph.text.replace(
                        "JEREMÍAS 45-46", "JEREMÍAS 45,46"
                    )
            row = next(
                row for row in document.tables[6].rows if row.cells[0].text == "7"
            )
            row.cells[2].text = "JORGE DANIEL VALLADARES(p)"
            document.tables[0].rows[1].cells[2].text = "MIRZA CERRANO"
            document.tables[0].rows[-1].cells[2].text = "SUSSY BANEGAS"
            document.save(source_path)

            weeks = parse_document(source_path)

        self.assertEqual(
            weeks[6].ministry[-1].first_participant_name(),
            "Jorge Daniel Valladares (p)",
        )
        self.assertEqual(
            weeks[0].bible_treasures[0].first_participant_name(), "Mirza Cerrano"
        )
        self.assertEqual(weeks[0].closing_prayer, "Sussy Banegas")

    def test_only_fully_uppercase_names_are_changed(self) -> None:
        self.assertEqual(normalize_person_name("JUAN PÉREZ"), "Juan Pérez")
        self.assertEqual(normalize_person_name("Juan McDONALD"), "Juan McDONALD")

    def test_suffix_is_preserved_when_uppercase_name_is_normalized(self) -> None:
        self.assertEqual(
            split_names("HELENA MARTÍNEZ (h) // CARLOS EJEMPLO (p)"),
            ("Helena Martínez (h)", "Carlos Ejemplo (p)"),
        )

    def test_dates_and_readings_use_sentence_case(self) -> None:
        self.assertEqual(
            normalize_sentence_case("7 AL 13 DE SEPTIEMBRE"),
            "7 al 13 de septiembre",
        )
        self.assertEqual(
            normalize_sentence_case("JEREMÍAS 32-33"),
            "Jeremías 32-33",
        )


if __name__ == "__main__":
    unittest.main()
