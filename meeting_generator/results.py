"""Resultado verificable del flujo completo de generación."""

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class GenerationResult:
    """Describe qué se persistió realmente durante una ejecución."""

    written_weeks: int
    omitted_weeks: int
    errors: tuple[str, ...]
    output_path: Path | None

    def __post_init__(self) -> None:
        if self.written_weeks < 0 or self.omitted_weeks < 0:
            raise ValueError("Los contadores de semanas no pueden ser negativos")
        if self.output_path is None:
            if self.written_weeks != 0 or not self.errors:
                raise ValueError("Un resultado sin archivo debe describir el fallo")
        elif self.errors or self.omitted_weeks != 0 or self.written_weeks == 0:
            raise ValueError("Un resultado con archivo debe representar éxito completo")

    @property
    def succeeded(self) -> bool:
        """Solo confirma éxito mientras el archivo persistido siga siendo válido."""
        if (
            self.output_path is None
            or self.errors
            or self.written_weeks <= 0
            or self.omitted_weeks != 0
        ):
            return False
        try:
            return self.output_path.is_file() and self.output_path.stat().st_size > 0
        except OSError:
            return False

    @property
    def requested_weeks(self) -> int:
        return self.written_weeks + self.omitted_weeks

    @classmethod
    def success(cls, written_weeks: int, output_path: Path) -> "GenerationResult":
        return cls(
            written_weeks=written_weeks,
            omitted_weeks=0,
            errors=(),
            output_path=output_path,
        )

    @classmethod
    def failure(
        cls,
        omitted_weeks: int,
        errors: tuple[str, ...] | list[str],
    ) -> "GenerationResult":
        normalized_errors = tuple(error.strip() for error in errors if error.strip())
        if not normalized_errors:
            normalized_errors = ("La generación falló sin un detalle disponible",)
        return cls(
            written_weeks=0,
            omitted_weeks=omitted_weeks,
            errors=normalized_errors,
            output_path=None,
        )
