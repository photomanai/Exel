#!/usr/bin/env python3
"""
Exel — Terminal-based, modern-looking CSV / spreadsheet editor.
Inspired by Lotus 1-2-3, but built with Textual (modern TUI framework)
as an "Excel alternative". Features mouse + keyboard support, formula support,
command palette, and theme support.
"""

import csv
import re
import string
import sys
from pathlib import Path
from typing import Optional

from textual.app import App, ComposeResult
from textual.containers import Horizontal, Vertical
from textual.widgets import (
    DataTable,
    Footer,
    Header,
    Input,
    Static,
    Label,
)
from textual.screen import ModalScreen
from textual.binding import Binding
from textual.coordinate import Coordinate
from textual import events


# --------------------------------------------------------------------------
# Helper functions: cell addressing (A1, B12, ...) and formula engine
# --------------------------------------------------------------------------

def col_letter(idx: int) -> str:
    """0 -> A, 1 -> B, ... 25 -> Z, 26 -> AA ..."""
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = string.ascii_uppercase[rem] + letters
    return letters


def col_index(letters: str) -> int:
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


CELL_RE = re.compile(r"([A-Za-z]+)(\d+)")
RANGE_RE = re.compile(r"([A-Za-z]+\d+):([A-Za-z]+\d+)")


class FormulaError(Exception):
    pass


class Sheet:
    """Simple formula-supported spreadsheet data model."""

    def __init__(self, rows: int = 60, cols: int = 26):
        self.n_rows = rows
        self.n_cols = cols
        self.data: dict[tuple[int, int], str] = {}
        self.filepath: Optional[Path] = None
        self.dirty = False

    # -- raw data access ---------------------------------------------------
    def raw(self, r: int, c: int) -> str:
        return self.data.get((r, c), "")

    def set_raw(self, r: int, c: int, value: str) -> None:
        if value == "":
            self.data.pop((r, c), None)
        else:
            self.data[(r, c)] = value
        self.dirty = True

    # -- formula evaluation ------------------------------------------------
    def value(self, r: int, c: int, _stack: Optional[set] = None) -> str:
        raw = self.raw(r, c)
        if not raw.startswith("="):
            return raw
        _stack = _stack or set()
        if (r, c) in _stack:
            return "#CIRCULAR!"
        _stack = _stack | {(r, c)}
        expr = raw[1:]
        try:
            return self._eval(expr, _stack)
        except FormulaError as e:
            return f"#ERR:{e}"
        except ZeroDivisionError:
            return "#DIV/0!"
        except Exception:
            return "#ERR"

    def _eval(self, expr: str, stack: set) -> str:
        expr = expr.strip()

        # If it's a single cell reference (=A1), return that cell's value
        # (even if it's text) as-is.
        single_ref = re.match(r"^([A-Za-z]+\d+)$", expr)
        if single_ref:
            rr, cc = self._parse_ref(single_ref.group(1))
            return self.value(rr, cc, stack)

        # Functions: SUM(A1:A5), AVERAGE(...), MIN(...), MAX(...), COUNT(...)
        func_match = re.match(
            r"^(SUM|AVERAGE|AVG|MIN|MAX|COUNT)\((.+)\)$", expr, re.IGNORECASE
        )
        if func_match:
            fname = func_match.group(1).upper()
            arg = func_match.group(2)
            nums = self._collect_numbers(arg, stack)
            if fname == "SUM":
                result = sum(nums)
            elif fname in ("AVERAGE", "AVG"):
                result = sum(nums) / len(nums) if nums else 0
            elif fname == "MIN":
                result = min(nums) if nums else 0
            elif fname == "MAX":
                result = max(nums) if nums else 0
            elif fname == "COUNT":
                result = len(nums)
            return self._fmt(result)

        # General arithmetic expression: replace cell references with numeric values
        def repl(m):
            ref = m.group(0)
            rr, cc = self._parse_ref(ref)
            v = self.value(rr, cc, stack)
            try:
                return str(float(v))
            except ValueError:
                return "0"

        safe_expr = CELL_RE.sub(repl, expr)
        if not re.match(r"^[0-9eE\.\+\-\*/\(\)\s]+$", safe_expr):
            raise FormulaError("invalid expression")
        try:
            result = eval(safe_expr, {"__builtins__": {}}, {})
        except Exception:
            raise FormulaError("cannot compute")
        return self._fmt(result)

    def _collect_numbers(self, arg: str, stack: set) -> list[float]:
        nums: list[float] = []
        for part in arg.split(","):
            part = part.strip()
            rm = RANGE_RE.match(part)
            if rm:
                r1, c1 = self._parse_ref(rm.group(1))
                r2, c2 = self._parse_ref(rm.group(2))
                for rr in range(min(r1, r2), max(r1, r2) + 1):
                    for cc in range(min(c1, c2), max(c1, c2) + 1):
                        v = self.value(rr, cc, stack)
                        try:
                            nums.append(float(v))
                        except ValueError:
                            pass
            elif CELL_RE.match(part):
                rr, cc = self._parse_ref(part)
                v = self.value(rr, cc, stack)
                try:
                    nums.append(float(v))
                except ValueError:
                    pass
            elif part:
                try:
                    nums.append(float(part))
                except ValueError:
                    pass
        return nums

    @staticmethod
    def _parse_ref(ref: str) -> tuple[int, int]:
        m = CELL_RE.match(ref.strip())
        if not m:
            raise FormulaError(f"invalid reference: {ref}")
        letters, digits = m.groups()
        return int(digits) - 1, col_index(letters)

    @staticmethod
    def _fmt(v: float) -> str:
        if float(v).is_integer():
            return str(int(v))
        return f"{v:.4g}"

    # -- CSV I/O -----------------------------------------------------------
    def load_csv(self, path: Path) -> None:
        self.data.clear()
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)
        self.n_rows = max(60, len(rows) + 5)
        self.n_cols = max(26, (max((len(r) for r in rows), default=0)) + 5)
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                if val != "":
                    self.data[(r, c)] = val
        self.filepath = path
        self.dirty = False

    def save_csv(self, path: Path) -> None:
        if not self.data:
            max_r, max_c = 0, 0
        else:
            max_r = max(r for r, _ in self.data) + 1
            max_c = max(c for _, c in self.data) + 1
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for r in range(max_r):
                writer.writerow([self.value(r, c) for c in range(max_c)])
        self.filepath = path
        self.dirty = False


# --------------------------------------------------------------------------
# Modal windows: file name prompt, save/open
# --------------------------------------------------------------------------

class PathPrompt(ModalScreen[Optional[str]]):
    """Simple file path input dialog (Open / Save As)."""

    DEFAULT_CSS = """
    PathPrompt {
        align: center middle;
    }
    #dialog {
        width: 60;
        height: auto;
        border: round $accent;
        background: $panel;
        padding: 1 2;
    }
    #dialog Label {
        margin-bottom: 1;
        text-style: bold;
    }
    """

    def __init__(self, title: str, default: str = ""):
        super().__init__()
        self.title_text = title
        self.default = default

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label(self.title_text)
            yield Input(value=self.default, placeholder="file.csv", id="path_input")

    def on_mount(self) -> None:
        self.query_one("#path_input", Input).focus()

    def on_input_submitted(self, event: Input.Submitted) -> None:
        self.dismiss(event.value.strip())

    def on_key(self, event: events.Key) -> None:
        if event.key == "escape":
            self.dismiss(None)


class ConfirmScreen(ModalScreen[bool]):
    DEFAULT_CSS = """
    ConfirmScreen {
        align: center middle;
    }
    #dialog {
        width: 50;
        height: auto;
        border: round $error;
        background: $panel;
        padding: 1 2;
    }
    """

    def __init__(self, message: str):
        super().__init__()
        self.message = message

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label(self.message)
            yield Label("[b]Yes[/b]: Enter   [b]No[/b]: Esc", classes="hint")

    def on_key(self, event: events.Key) -> None:
        if event.key == "enter":
            self.dismiss(True)
        elif event.key == "escape":
            self.dismiss(False)


# --------------------------------------------------------------------------
# Main application
# --------------------------------------------------------------------------

class Exel(App):
    TITLE = "Exel — Terminal Spreadsheet"
    SUB_TITLE = "new file"

    CSS = """
    Screen {
        background: $surface;
    }
    #formula_bar {
        height: 3;
        background: $panel;
        border: round $accent;
        padding: 0 1;
    }
    #cell_addr {
        width: 8;
        color: $accent;
        text-style: bold;
        content-align: center middle;
    }
    #formula_input {
        border: none;
    }
    #status_bar {
        height: 1;
        background: $accent 20%;
        color: $text;
        padding: 0 1;
    }
    DataTable {
        height: 1fr;
    }
    .hint {
        color: $text-muted;
    }
    """

    BINDINGS = [
        Binding("ctrl+s", "save", "Save"),
        Binding("ctrl+shift+s", "save_as", "Save As"),
        Binding("ctrl+o", "open", "Open"),
        Binding("ctrl+n", "new_sheet", "New"),
        Binding("enter", "edit_cell", "Edit", show=True),
        Binding("delete", "clear_cell", "Clear"),
        Binding("backspace", "clear_cell", "Clear", show=False),
        Binding("ctrl+r", "insert_row", "Insert Row"),
        Binding("ctrl+k", "insert_col", "Insert Column"),
        Binding("ctrl+d", "delete_row", "Delete Row"),
        Binding("f1", "help", "Help"),
        Binding("ctrl+q", "quit", "Quit"),
    ]

    def __init__(self, initial_file: Optional[str] = None):
        super().__init__()
        self.sheet = Sheet()
        self.initial_file = initial_file
        self.editing = False

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        with Horizontal(id="formula_bar"):
            yield Label("A1", id="cell_addr")
            yield Input(id="formula_input", placeholder="cell content or =FORMULA(...)")
        yield DataTable(id="grid", zebra_stripes=True, cursor_type="cell")
        yield Static("Ready  |  Enter: edit · Ctrl+S: save · Ctrl+O: open · F1: help", id="status_bar")
        yield Footer()

    def on_mount(self) -> None:
        table = self.query_one(DataTable)
        table.add_column("", key="rownum", width=5)
        for c in range(self.sheet.n_cols):
            table.add_column(col_letter(c), key=f"c{c}", width=10)
        for r in range(self.sheet.n_rows):
            row_cells = [str(r + 1)] + ["" for _ in range(self.sheet.n_cols)]
            table.add_row(*row_cells, key=f"r{r}")
        table.cursor_type = "cell"
        table.cursor_coordinate = Coordinate(0, 1)  # start at cell A1
        table.focus()
        self.update_formula_bar()

        if self.initial_file:
            self._load_file(self.initial_file)

    # -- helpers -----------------------------------------------------------
    def current_rc(self) -> tuple[int, int]:
        table = self.query_one(DataTable)
        coord = table.cursor_coordinate
        # subtract 1 from column because first column is row numbers
        return coord.row, coord.column - 1

    def refresh_cell_display(self, r: int, c: int) -> None:
        table = self.query_one(DataTable)
        val = self.sheet.value(r, c)
        table.update_cell(f"r{r}", f"c{c}", val)

    def update_formula_bar(self) -> None:
        r, c = self.current_rc()
        if c < 0:
            return
        addr = f"{col_letter(c)}{r + 1}"
        self.query_one("#cell_addr", Label).update(addr)
        finput = self.query_one("#formula_input", Input)
        finput.value = self.sheet.raw(r, c)

    def on_data_table_cell_highlighted(self, event: DataTable.CellHighlighted) -> None:
        self.update_formula_bar()

    def on_data_table_cell_selected(self, event: DataTable.CellSelected) -> None:
        # DataTable binds "enter" to its own select_cursor action;
        # we catch it here to start cell editing.
        self.action_edit_cell()

    def set_status(self, msg: str) -> None:
        self.query_one("#status_bar", Static).update(msg)

    # -- editing -----------------------------------------------------------
    def action_edit_cell(self) -> None:
        finput = self.query_one("#formula_input", Input)
        finput.focus()
        finput.selection_anchor = None
        finput.cursor_position = len(finput.value)

    def on_input_submitted(self, event: Input.Submitted) -> None:
        if event.input.id == "formula_input":
            r, c = self.current_rc()
            if c >= 0:
                self.sheet.set_raw(r, c, event.value)
                self.refresh_cell_display(r, c)
                self._recalc_all()
                self.set_status(f"{col_letter(c)}{r+1} updated")
            table = self.query_one(DataTable)
            table.focus()

    def action_clear_cell(self) -> None:
        table = self.query_one(DataTable)
        if table.has_focus:
            r, c = self.current_rc()
            if c >= 0:
                self.sheet.set_raw(r, c, "")
                self.refresh_cell_display(r, c)
                self.update_formula_bar()

    def _recalc_all(self) -> None:
        # Simple approach: redraw all cells that contain formulas
        for (r, c), raw in list(self.sheet.data.items()):
            if raw.startswith("="):
                self.refresh_cell_display(r, c)

    # -- row / column operations -------------------------------------------
    def action_insert_row(self) -> None:
        table = self.query_one(DataTable)
        r, _ = self.current_rc()
        # shift data down
        new_data = {}
        for (rr, cc), v in self.sheet.data.items():
            new_data[(rr + 1, cc)] if rr >= r else None
        shifted = {}
        for (rr, cc), v in self.sheet.data.items():
            shifted[(rr + 1 if rr >= r else rr, cc)] = v
        self.sheet.data = shifted
        self.sheet.n_rows += 1
        table.add_row(*([str(self.sheet.n_rows)] + [""] * self.sheet.n_cols), key=f"r{self.sheet.n_rows-1}")
        self._redraw_all()
        self.set_status(f"New row inserted before row {r+1}")

    def action_delete_row(self) -> None:
        r, _ = self.current_rc()
        shifted = {}
        for (rr, cc), v in self.sheet.data.items():
            if rr == r:
                continue
            shifted[(rr - 1 if rr > r else rr, cc)] = v
        self.sheet.data = shifted
        self._redraw_all()
        self.set_status(f"Row {r+1} deleted")

    def action_insert_col(self) -> None:
        _, c = self.current_rc()
        if c < 0:
            c = 0
        shifted = {}
        for (rr, cc), v in self.sheet.data.items():
            shifted[(rr, cc + 1 if cc >= c else cc)] = v
        self.sheet.data = shifted
        self.sheet.n_cols += 1
        table = self.query_one(DataTable)
        table.add_column(col_letter(self.sheet.n_cols - 1), key=f"c{self.sheet.n_cols-1}", width=10)
        self._redraw_all()
        self.set_status(f"New column inserted before column {col_letter(c)}")

    def _redraw_all(self) -> None:
        table = self.query_one(DataTable)
        for r in range(self.sheet.n_rows):
            for c in range(self.sheet.n_cols):
                val = self.sheet.value(r, c)
                try:
                    table.update_cell(f"r{r}", f"c{c}", val)
                except Exception:
                    pass
        self.update_formula_bar()

    # -- file operations ---------------------------------------------------
    def _load_file(self, path_str: str) -> None:
        path = Path(path_str).expanduser()
        try:
            if path.exists():
                self.sheet.load_csv(path)
                self._rebuild_table()
                self.set_status(f"Loaded: {path}")
            else:
                self.sheet.filepath = path
                self.set_status(f"New file: {path}")
            self.sub_title = str(path)
        except Exception as e:
            self.set_status(f"Error: {e}")

    def _rebuild_table(self) -> None:
        table = self.query_one(DataTable)
        table.clear(columns=True)
        table.add_column("", key="rownum", width=5)
        for c in range(self.sheet.n_cols):
            table.add_column(col_letter(c), key=f"c{c}", width=10)
        for r in range(self.sheet.n_rows):
            row_cells = [str(r + 1)] + [self.sheet.value(r, c) for c in range(self.sheet.n_cols)]
            table.add_row(*row_cells, key=f"r{r}")
        self.update_formula_bar()

    def action_save(self) -> None:
        if self.sheet.filepath:
            self.sheet.save_csv(self.sheet.filepath)
            self.set_status(f"Saved: {self.sheet.filepath}")
        else:
            self.action_save_as()

    def action_save_as(self) -> None:
        def cb(result: Optional[str]) -> None:
            if result:
                path = Path(result).expanduser()
                if not path.suffix:
                    path = path.with_suffix(".csv")
                self.sheet.save_csv(path)
                self.sub_title = str(path)
                self.set_status(f"Saved: {path}")

        self.push_screen(PathPrompt("Save as — file path:", str(self.sheet.filepath or "sheet.csv")), cb)

    def action_open(self) -> None:
        def cb(result: Optional[str]) -> None:
            if result:
                self._load_file(result)

        self.push_screen(PathPrompt("Open file — file path:"), cb)

    def action_new_sheet(self) -> None:
        self.sheet = Sheet()
        self._rebuild_table()
        self.sub_title = "new file"
        self.set_status("New blank sheet")

    def action_help(self) -> None:
        self.set_status(
            "Enter:edit  Ctrl+S:save  Ctrl+Shift+S:save as  "
            "Ctrl+O:open  Ctrl+N:new  Ctrl+R:insert row  Ctrl+K:insert column  "
            "Ctrl+D:delete row  Del:clear  Formula: =SUM(A1:A5) =AVERAGE(...) =A1+B1*2"
        )


def main():
    initial = sys.argv[1] if len(sys.argv) > 1 else None
    app = Exel(initial_file=initial)
    app.run()


if __name__ == "__main__":
    main()
