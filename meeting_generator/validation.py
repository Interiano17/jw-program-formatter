"""Errores de validación que impiden generar un documento incompleto."""

from collections.abc import Iterable


class GenerationValidationError(ValueError):
    """Agrupa todos los problemas detectados antes de guardar el resultado."""

    def __init__(self, context: str, issues: Iterable[str]) -> None:
        normalized_issues = tuple(issue.strip() for issue in issues if issue.strip())
        if not normalized_issues:
            raise ValueError("GenerationValidationError requiere al menos un problema")

        self.context = context
        self.issues = normalized_issues
        details = "\n".join(f"- {issue}" for issue in normalized_issues)
        super().__init__(f"{context}:\n{details}")
