"""Panel desplegable que muestra los errores de validación agrupados por semana."""

import re
import tkinter as tk
from collections.abc import Sequence
from tkinter import font as tkfont
from tkinter import ttk

from .theme import (
    COLOR_ACCENT,
    COLOR_BORDER,
    COLOR_CARD,
    COLOR_ERROR,
    COLOR_INFO,
    COLOR_PRIMARY,
)

GENERAL_GROUP = "General"
_WEEK_ERROR = re.compile(
    r"^semana\s+(\d+)\s*(?:,\s*([^:]+?))?\s*:\s*(.+)$",
    re.IGNORECASE | re.DOTALL,
)


def group_errors(errors: Sequence[str]) -> list[tuple[str, list[str]]]:
    """Agrupa mensajes "semana N: ..." por semana; el resto va en General."""
    general: list[str] = []
    by_week: dict[int, list[str]] = {}
    for error in errors:
        text = error.strip()
        if not text:
            continue
        match = _WEEK_ERROR.match(text)
        if match is None:
            general.append(text)
            continue
        week = int(match.group(1))
        detail = match.group(3).strip()
        if match.group(2):
            detail = f"{match.group(2).strip()}: {detail}"
        by_week.setdefault(week, []).append(detail[:1].upper() + detail[1:])

    groups: list[tuple[str, list[str]]] = []
    if general:
        groups.append((GENERAL_GROUP, general))
    groups.extend((f"Semana {week}", by_week[week]) for week in sorted(by_week))
    return groups


class ErrorPanel(tk.Frame):
    """Tarjeta con texto desplazable y botón para copiar los errores."""

    def __init__(self, parent: tk.Misc, *, background: str, family: str) -> None:
        super().__init__(parent, bg=background)
        self._background = background
        self._plain_text = ""

        header = tk.Frame(self, bg=background)
        header.pack(fill=tk.X, pady=(0, 6))
        self._title = tk.Label(
            header,
            text="",
            font=tkfont.Font(family=family, size=8, weight="bold"),
            bg=background,
            fg=COLOR_ERROR,
            anchor="w",
        )
        self._title.pack(side=tk.LEFT)
        ttk.Button(
            header,
            text="Copiar",
            style="Compact.Secondary.TButton",
            command=self._copy,
            cursor="hand2",
        ).pack(side=tk.RIGHT)

        body = tk.Frame(
            self,
            bg=COLOR_CARD,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_BORDER,
            highlightthickness=1,
        )
        body.pack(fill=tk.BOTH, expand=True)
        self._text = tk.Text(
            body,
            height=10,
            wrap="word",
            relief="flat",
            borderwidth=0,
            padx=14,
            pady=10,
            bg=COLOR_CARD,
            fg=COLOR_INFO,
            font=tkfont.Font(family=family, size=10),
            cursor="arrow",
            state=tk.DISABLED,
            spacing3=3,
        )
        scrollbar = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self._text.yview)
        self._text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._text.tag_configure(
            "group",
            font=tkfont.Font(family=family, size=10, weight="bold"),
            foreground=COLOR_ACCENT,
            spacing1=8,
        )
        self._text.tag_configure(
            "item", lmargin1=12, lmargin2=24, foreground=COLOR_INFO
        )
        self._text.tag_configure("bullet", foreground=COLOR_PRIMARY)

    def show(self, errors: Sequence[str]) -> None:
        """Muestra el panel con los errores agrupados."""
        groups = group_errors(errors)
        total = sum(len(items) for _, items in groups)
        self._title.config(
            text=f"DETALLES · {total} PROBLEMA{'S' if total != 1 else ''}"
        )

        lines: list[str] = []
        self._text.config(state=tk.NORMAL)
        self._text.delete("1.0", tk.END)
        for title, items in groups:
            self._text.insert(tk.END, f"{title}\n", "group")
            lines.append(title)
            for item in items:
                self._text.insert(tk.END, "•  ", ("item", "bullet"))
                self._text.insert(tk.END, f"{item}\n", "item")
                lines.append(f"  - {item}")
        self._text.config(state=tk.DISABLED)
        self._text.yview_moveto(0)
        self._plain_text = "\n".join(lines)
        if not self.winfo_ismapped():
            self.pack(fill=tk.BOTH, expand=True, pady=(16, 0))

    def hide(self) -> None:
        """Oculta el panel y descarta su contenido."""
        self._plain_text = ""
        self.pack_forget()

    @property
    def visible(self) -> bool:
        return bool(self.winfo_manager())

    def _copy(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self._plain_text)
