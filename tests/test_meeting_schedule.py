import copy
import os
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from docx import Document

from meeting_generator.config import BIBLE_BOOK_NAMES
from meeting_generator.generator import generate_document
from meeting_generator.main import MeetingGeneratorApp
from meeting_generator.models import Assignment
from meeting_generator.parser import ProgramParser, parse_document
from meeting_generator.results import GenerationResult
from meeting_generator.template_writer import TemplateWriter, fill_template
from meeting_generator.utils import normalize_person_name, normalize_sentence_case
from meeting_generator.validation import GenerationValidationError
from tests.source_fixtures import build_primary_source, build_secondary_source
from tests.template_fixture import build_template

_FIXTURE_DIRECTORY: tempfile.TemporaryDirectory[str] | None = None
_CONGREGATION_PATCHER = None
SOURCE_DOCUMENT = Path("primary-source-not-built.docx")
SECOND_SOURCE_DOCUMENT = Path("secondary-source-not-built.docx")
TEMPLATE_DOCUMENT = Path("template-not-built.docx")


def setUpModule() -> None:
    """Construye todos los DOCX de prueba sin depender del árbol local."""
    global _CONGREGATION_PATCHER
    global _FIXTURE_DIRECTORY
    global SECOND_SOURCE_DOCUMENT
    global SOURCE_DOCUMENT
    global TEMPLATE_DOCUMENT

    _FIXTURE_DIRECTORY = tempfile.TemporaryDirectory()
    fixture_directory = Path(_FIXTURE_DIRECTORY.name)
    SOURCE_DOCUMENT = build_primary_source(
        fixture_directory / "primary-source.docx"
    )
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

    def test_intermediate_and_closing_songs_are_distinct(self) -> None:
        first_week = self.weeks[0]

        self.assertEqual(first_week.intermediate_song, "49")
        self.assertEqual(first_week.closing_song, "61")


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
        header = self._extract_header(
            "SEMANA DEL 5 AL 11 DE ENERO. ISAÍAS 1-3"
        )

        self.assertEqual(header["date"], "5 al 11 de enero")
        self.assertEqual(header["weekly_reading"], "Isaías 1-3")

    def test_numbered_book_preserves_ordinal(self) -> None:
        header = self._extract_header(
            "SEMANA DEL 12 AL 18 DE ENERO. 1 CORINTIOS 1-2"
        )

        self.assertEqual(header["date"], "12 al 18 de enero")
        self.assertEqual(header["weekly_reading"], "1 Corintios 1-2")

    def test_oseas_is_recognized_as_weekly_reading(self) -> None:
        header = self._extract_header(
            "SEMANA DEL 19 AL 25 DE ENERO. OSEAS 1-3"
        )

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
                    "SEMANA DEL 26 DE ENERO AL 1 DE FEBRERO. "
                    f"{source_reading}"
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
        weeks = parse_document(SOURCE_DOCUMENT)
        fill_template(TEMPLATE_DOCUMENT, weeks, output_path)
        cls.document = Document(output_path)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.temp_directory.cleanup()

    def test_first_week_uses_expected_start_times(self) -> None:
        first_copy = self.document.tables[0].rows[:27]
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
        rows = self.document.tables[2].rows

        self.assertEqual(rows[46].cells[0].text.strip(), "7:49")
        self.assertIn("(10 mins.)", rows[46].cells[2].text)
        self.assertEqual(rows[47].cells[0].text.strip(), "7:59")
        self.assertIn("(5 mins.)", rows[47].cells[2].text)
        self.assertEqual(rows[48].cells[0].text.strip(), "8:04")

    def test_generated_document_has_no_zero_time_markers(self) -> None:
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    self.assertNotIn("0:00", cell.text)


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
        first_copy = self.document.tables[0].rows[:27]
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
                "7:43",
                "7:58",
                "8:28",
                "8:31",
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
            any("semana 1" in issue.lower() for issue in raised.exception.issues),
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
        self.app.root = Mock()
        self.app.status_label = Mock()
        self.app.source_button = Mock()
        self.app.template_button = Mock()
        self.app.generate_button = Mock()

    def tearDown(self) -> None:
        self.temp_directory.cleanup()

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
            self.app._generate_document()

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
            self.app._generate_document()

        show_info.assert_not_called()
        show_error.assert_called_once()
        failure_message = show_error.call_args.args[1]
        self.assertIn("4", failure_message)
        self.assertIn("capacidad", failure_message)

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
            self.app._generate_document()

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
            self.app._generate_document()

        show_error.assert_not_called()
        generate.assert_called_once_with(
            source_snapshot,
            template_snapshot,
            self.output_path,
        )

    def test_reentrant_generation_is_ignored(self) -> None:
        self.app._generation_in_progress = True

        with (
            patch(
                "meeting_generator.main.filedialog.asksaveasfilename"
            ) as save_as,
            patch("meeting_generator.main.generate_document") as generate,
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self.app._generate_document()

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
            self.app._generate_document()

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
            self.app._generate_document()

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
                self.app._generate_document()

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
            patch("meeting_generator.main.messagebox.showinfo") as show_info,
            patch("meeting_generator.main.messagebox.showerror") as show_error,
        ):
            self.app._generate_document()

        show_info.assert_not_called()
        show_error.assert_called_once()
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


class NameNormalizationTests(unittest.TestCase):
    def test_only_fully_uppercase_names_are_changed(self) -> None:
        self.assertEqual(normalize_person_name("JUAN PÉREZ"), "Juan Pérez")
        self.assertEqual(normalize_person_name("Juan McDONALD"), "Juan McDONALD")

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
