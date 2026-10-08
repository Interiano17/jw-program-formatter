import io
import logging
import tempfile
import unittest
from logging.handlers import RotatingFileHandler
from pathlib import Path
from unittest.mock import patch

from docx import Document

from meeting_generator.config import LOG_BACKUP_COUNT, LOG_MAX_BYTES
from meeting_generator.generator import generate_document
from meeting_generator.main import main
from meeting_generator.parser import ProgramParser, parse_document
from meeting_generator.utils import setup_logging
from tests.source_fixtures import build_secondary_source
from tests.template_fixture import build_template


class LoggingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.folder = Path(self.directory.name)
        self.log_path = self.folder / "application.log"
        self.console = io.StringIO()
        self.stdout_patch = patch("meeting_generator.utils.sys.stdout", self.console)
        self.stdout_patch.start()
        self.logger = logging.getLogger("meeting_generator")
        self.previous_handlers = self.logger.handlers[:]
        self.previous_level = self.logger.level
        self.previous_propagate = self.logger.propagate
        for handler in self.previous_handlers:
            self.logger.removeHandler(handler)

    def tearDown(self) -> None:
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)
            handler.close()
        for handler in self.previous_handlers:
            self.logger.addHandler(handler)
        self.logger.setLevel(self.previous_level)
        self.logger.propagate = self.previous_propagate
        self.stdout_patch.stop()
        self.directory.cleanup()

    def test_repeated_setup_writes_once_and_preserves_external_handlers(self) -> None:
        root = logging.getLogger()
        root_level = root.level
        root_console = io.StringIO()
        external_root = logging.StreamHandler(root_console)
        external_package = logging.StreamHandler(io.StringIO())
        root.addHandler(external_root)
        self.addCleanup(root.removeHandler, external_root)
        self.addCleanup(external_root.close)
        self.logger.addHandler(external_package)
        root_handlers = root.handlers[:]

        self.assertTrue(setup_logging(self.log_path))
        old_file = next(
            handler
            for handler in self.logger.handlers
            if isinstance(handler, RotatingFileHandler)
        )
        setup_logging(self.log_path)
        setup_logging(self.log_path)
        logging.getLogger("meeting_generator.parser").info("registro único")
        root.error("registro externo")

        self.assertEqual(root.handlers, root_handlers)
        self.assertEqual(root.level, root_level)
        self.assertIn(external_package, self.logger.handlers)
        self.assertEqual(len(self.logger.handlers), 3)
        self.assertIsNone(old_file.stream)
        self.assertEqual(self.console.getvalue().count("registro único"), 1)
        contents = self.log_path.read_text(encoding="utf-8")
        self.assertEqual(contents.count("registro único"), 1)
        self.assertNotIn("registro externo", contents)
        self.assertNotIn("registro único", root_console.getvalue())
        self.assertIn("registro externo", root_console.getvalue())
        current_file = next(
            handler
            for handler in self.logger.handlers
            if isinstance(handler, RotatingFileHandler)
        )
        self.assertEqual(current_file.maxBytes, LOG_MAX_BYTES)
        self.assertEqual(current_file.backupCount, LOG_BACKUP_COUNT)

    def test_new_path_closes_previous_file_and_redirects_records(self) -> None:
        setup_logging(self.log_path)
        previous_file = next(
            handler
            for handler in self.logger.handlers
            if isinstance(handler, RotatingFileHandler)
        )
        self.logger.info("registro previo")
        new_path = self.folder / "new.log"

        setup_logging(new_path)
        self.logger.info("registro posterior")

        self.assertIsNone(previous_file.stream)
        self.assertEqual(len(self.logger.handlers), 2)
        self.assertIn("registro previo", self.log_path.read_text(encoding="utf-8"))
        self.assertNotIn(
            "registro posterior", self.log_path.read_text(encoding="utf-8")
        )
        self.assertIn("registro posterior", new_path.read_text(encoding="utf-8"))

    def test_failed_setup_preserves_working_handlers(self) -> None:
        setup_logging(self.log_path)
        previous_handlers = self.logger.handlers[:]
        previous_level = self.logger.level
        previous_propagate = self.logger.propagate

        self.assertFalse(setup_logging(self.folder / "missing" / "application.log"))
        self.logger.info("configuración conservada")

        self.assertEqual(self.logger.handlers, previous_handlers)
        self.assertEqual(self.logger.level, previous_level)
        self.assertEqual(self.logger.propagate, previous_propagate)
        self.assertEqual(self.console.getvalue().count("configuración conservada"), 1)
        self.assertIn(
            "configuración conservada", self.log_path.read_text(encoding="utf-8")
        )

    def test_unavailable_log_paths_keep_console_logging_without_duplicates(
        self,
    ) -> None:
        for log_path in (self.folder / "missing" / "application.log", self.folder):
            with self.subTest(log_path=log_path):
                self.console.seek(0)
                self.console.truncate(0)

                self.assertFalse(setup_logging(log_path))
                self.assertFalse(setup_logging(log_path))
                logging.getLogger("meeting_generator.parser").info(
                    "registro disponible"
                )

                self.assertEqual(len(self.logger.handlers), 1)
                self.assertEqual(
                    self.logger.handlers[0].name, "meeting_generator.console"
                )
                self.assertFalse(self.logger.propagate)
                contents = self.console.getvalue()
                self.assertEqual(contents.count("registro disponible"), 1)
                self.assertIn("WARNING", contents)
                self.assertIn("No se pudo abrir el archivo de log", contents)
                self.assertIn("consola", contents)
                self.assertNotIn(str(log_path), contents)
                self.assertNotIn("application.log", contents)

    def test_permission_error_keeps_console_without_exposing_error_details(
        self,
    ) -> None:
        with patch(
            "meeting_generator.utils.RotatingFileHandler",
            side_effect=PermissionError(f"DirectorioReservado: {self.log_path}"),
        ):
            self.assertFalse(setup_logging(self.log_path))
        self.logger.info("registro sin archivo")

        contents = self.console.getvalue()
        self.assertIn("registro sin archivo", contents)
        self.assertIn("No se pudo abrir el archivo de log", contents)
        self.assertNotIn("DirectorioReservado", contents)
        self.assertNotIn(str(self.log_path), contents)

    def test_missing_current_directory_keeps_console_logging(self) -> None:
        with patch(
            "meeting_generator.utils.Path.cwd",
            side_effect=FileNotFoundError("DirectorioReservado"),
        ):
            self.assertFalse(setup_logging())
        self.logger.info("registro sin directorio")

        self.assertIn("registro sin directorio", self.console.getvalue())
        self.assertIn("No se pudo abrir el archivo de log", self.console.getvalue())
        self.assertNotIn("DirectorioReservado", self.console.getvalue())
        self.assertFalse(self.logger.propagate)

    def test_file_logging_can_recover_after_console_fallback(self) -> None:
        self.assertFalse(setup_logging(self.folder / "missing" / "application.log"))
        previous_console = self.logger.handlers[0]

        with patch.object(
            previous_console, "close", wraps=previous_console.close
        ) as close:
            self.assertTrue(setup_logging(self.log_path))

        close.assert_called_once_with()
        self.assertNotIn(previous_console, self.logger.handlers)
        self.assertEqual(len(self.logger.handlers), 2)
        self.logger.info("registro recuperado")
        self.assertEqual(self.console.getvalue().count("registro recuperado"), 1)
        self.assertEqual(
            self.log_path.read_text(encoding="utf-8").count("registro recuperado"), 1
        )

    def test_application_starts_and_warns_when_first_log_open_fails(self) -> None:
        with (
            patch(
                "meeting_generator.utils.Path.cwd", return_value=self.folder / "missing"
            ),
            patch("meeting_generator.main.MeetingGeneratorApp") as application,
            patch("meeting_generator.main.messagebox.showwarning") as warning,
        ):
            main()

        application.assert_called_once_with()
        warning.assert_called_once()
        self.assertEqual(
            warning.call_args.kwargs, {"parent": application.return_value.root}
        )
        self.assertIn("consola", warning.call_args.args[1])
        self.assertNotIn(str(self.folder), warning.call_args.args[1])
        application.return_value.run.assert_called_once_with()

    def test_application_starts_without_warning_when_log_is_available(self) -> None:
        with (
            patch("meeting_generator.utils.Path.cwd", return_value=self.folder),
            patch("meeting_generator.main.MeetingGeneratorApp") as application,
            patch("meeting_generator.main.messagebox.showwarning") as warning,
        ):
            main()

        application.assert_called_once_with()
        warning.assert_not_called()
        application.return_value.run.assert_called_once_with()
        self.assertTrue(
            any(
                isinstance(handler, RotatingFileHandler)
                for handler in self.logger.handlers
            )
        )

    def test_rotation_limits_size_and_retained_backups(self) -> None:
        with (
            patch("meeting_generator.utils.LOG_MAX_BYTES", 512),
            patch("meeting_generator.utils.LOG_BACKUP_COUNT", 2),
        ):
            setup_logging(self.log_path)
            for index in range(30):
                self.logger.info("rotation-record-%02d %s", index, "x" * 90)

        log_files = list(self.folder.glob("application.log*"))
        self.assertEqual(
            {path.name for path in log_files},
            {"application.log", "application.log.1", "application.log.2"},
        )
        self.assertTrue(all(0 < path.stat().st_size <= 512 for path in log_files))
        self.assertIn("rotation-record-29", self.log_path.read_text(encoding="utf-8"))
        self.assertTrue(
            all(
                "rotation-record-00" not in path.read_text(encoding="utf-8")
                for path in log_files
            )
        )

    def test_routine_and_debug_logs_omit_personal_data_and_source_content(self) -> None:
        private_folder = self.folder / "DirectorioPersonalReservado"
        private_folder.mkdir()
        source = build_secondary_source(private_folder / "FuentePersonalReservada.docx")
        template = build_template(private_folder / "PlantillaPersonalReservada.docx")
        output = private_folder / "SalidaPersonalReservada.docx"
        document = Document(source)
        content_token = "ContenidoPrivadoReservado"
        for table in document.tables:
            for row in table.rows:
                cells = row.cells
                if cells[1].text.startswith(("DISCURSO", "Tema sintético")):
                    cells[1].text += f" {content_token}"
        document.save(source)
        weeks = parse_document(source)
        personal_names = {
            participant.name
            for week in weeks
            for assignment in week.all_assignments()
            for participant in assignment.participants
        } | {week.president for week in weeks}

        with patch("meeting_generator.utils.LOG_LEVEL", logging.DEBUG):
            setup_logging(self.log_path)
            result = generate_document(source, template, output)
            ProgramParser(source)._create_assignment(99, ["5 mins."], "")

        self.assertTrue(result.succeeded, result.errors)
        for contents in (
            self.log_path.read_text(encoding="utf-8"),
            self.console.getvalue(),
        ):
            for private_value in personal_names | {
                private_folder.name,
                source.name,
                template.name,
                output.name,
                content_token,
            }:
                self.assertNotIn(private_value, contents)
            self.assertIn("Semana 1:", contents)
            self.assertIn("número 8 inferido", contents)
            self.assertIn("sin duración detectada", contents)
            self.assertIn("Punto 99: título vacío", contents)
            self.assertIn("Resultado: escritas=8 omitidas=0 errores=0", contents)


if __name__ == "__main__":
    unittest.main()
