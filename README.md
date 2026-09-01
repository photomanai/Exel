# Exel — Modern Terminal Spreadsheet

A CSV editor / simple Excel alternative inspired by Lotus 1‑2‑3 but built with
**Textual** (modern TUI framework), running in the terminal with mouse + keyboard
support, delivering a true "terminal GUI" feel.

## Features

- 📊 Excel-like cell grid (row/column headers: A, B, C… / 1, 2, 3…)
- 🧮 Formula support: `=A1+B2`, `=SUM(A1:A5)`, `=AVERAGE(A1:B3)`, `=MIN(...)`,
  `=MAX(...)`, `=COUNT(...)`, cell references (`=A1`) and nested arithmetic
- 💾 CSV open / save / save as
- ➕ Insert and delete rows/columns
- 🖱️ Mouse support (click cells, column width, etc.)
- 🎨 Dark theme, formula bar, status bar
- ⌨️ Keyboard shortcuts + command palette (Textual's built-in `Ctrl+P` palette)

## Installation

```bash
pip install -r requirements.txt
# or single line:
pip install textual
```

## Running

```bash
python3 app.py                 # start with a blank sheet
python3 app.py data.csv        # start by opening an existing CSV file
```

## Keyboard Shortcuts

| Key | Action |
|---|---|
| `Arrow keys` / mouse | Navigate between cells |
| `Enter` | Edit selected cell (focus formula bar) |
| `Delete` / `Backspace` | Clear cell |
| `Ctrl+S` | Save |
| `Ctrl+Shift+S` | Save as |
| `Ctrl+O` | Open file |
| `Ctrl+N` | New blank sheet |
| `Ctrl+R` | Insert new row above selected row |
| `Ctrl+K` | Insert new column before selected column |
| `Ctrl+D` | Delete selected row |
| `F1` | Show shortcut help in status bar |
| `Ctrl+Q` | Quit |
| `Ctrl+P` | Command palette (including theme switching) |

## Formula Examples

```
=10+5*2
=A1+B1
=SUM(A1:A10)
=AVERAGE(B2:B20)
=MAX(C1:C5)
=A1        (returns the content/text of another cell as-is)
```

> Note: CSV format does not store formulas; when you save, the
> **computed values** are written to the file (think of it like
> "paste values" in Excel).

## File Structure

```
lotus3k/
├── app.py            # Full application (single file, ~450 lines)
├── requirements.txt
└── README.md
```

## Extension Ideas

- Cell formatting (color, bold, currency format)
- Multi-sheet support and tabs
- Undo/redo
- Real-time multi-user editing (via WebSocket)
