# Changelog

All notable changes to this project will be documented here. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] — 2026-04-17

First release. Pre-alpha — API may evolve.

### Added

- `ExVEx.open/1` and `ExVEx.save/2` — byte-preserving round-trip on every
  part in the archive that the caller has not explicitly mutated, including
  `.xlsm` VBA binaries, custom XML, and unknown content types.
- `ExVEx.sheet_names/1` and `ExVEx.sheet_path/2` for sheet navigation.
- `ExVEx.get_cell/3` — reads strings (shared & inline), numbers, booleans,
  dates (as `Date`), date-times (as `NaiveDateTime`), formula results, and
  cell errors.
- `ExVEx.put_cell/4` — writes:
  - strings (deduplicated through the shared-string table when present; falls
    back to inline strings when the workbook has none)
  - numbers (integers and floats)
  - booleans
  - `nil` (clears the cell)
  - `{:formula, "..."}` and `{:formula, "...", cached_value}`
  - `Date` and `NaiveDateTime` (converted to Excel serial numbers; an xf
    with the matching date numFmtId is added to `xl/styles.xml` if one
    isn't already present)
- `ExVEx.get_formula/3` — reads the formula string from a formula cell.
- `ExVEx.get_style/3` — resolves a cell's style into a flat
  `%ExVEx.Style{}` with font, fill, border, alignment, and number-format
  sub-records dereferenced from the stylesheet.
- `ExVEx.cells/2` — returns every populated cell on a sheet as a
  `%{ref => value}` map.
- `ExVEx.each_cell/2` — streams every populated cell in row-major order.
- `ExVEx.merge_cells/3,4`, `ExVEx.unmerge_cells/3,4`,
  `ExVEx.merged_ranges/2` — merged-cell management with Excel-faithful
  defaults (clears non-anchor cells on merge; exact-match required on
  unmerge). Options: `:preserve_values` (`false` | `true`),
  `:on_overlap` (`:error` | `:replace` | `:allow`), `:on_missing`
  (`:error` | `:ignore`). Ranges are stored as
  `<mergeCells><mergeCell ref="A1:B2"/></mergeCells>` in the worksheet
  XML, inserted in the correct position in the schema's element order.
- OOXML parsers: `Packaging.ContentTypes`, `Packaging.Relationships`,
  `OOXML.Workbook`, `OOXML.Worksheet`, `OOXML.SharedStrings`, `OOXML.Styles`.
- Style model: `ExVEx.Style` + `Font`, `Fill`, `Border`, `Side`,
  `Alignment`, `Color`.
- Coordinate utilities: `ExVEx.Utils.Coordinate` — A1 ↔ `{row, col}`,
  Excel's bijective base-26 column labels.

### Formula freshness on save

When a workbook is mutated, ExVEx invalidates the calculation chain cache
so Excel recomputes formulas on open instead of showing stale `#N/A`
placeholders. On save of a dirty workbook:

- `xl/calcChain.xml` is dropped from the archive.
- Its entry is removed from `[Content_Types].xml` and
  `xl/_rels/workbook.xml.rels`.
- `<calcPr fullCalcOnLoad="1">` is set on `xl/workbook.xml`.

No-op saves (open → save without mutation) leave every part byte-identical.

### Coordinate addressing

Every cell-addressing function (`get_cell/3`, `put_cell/4`, `get_formula/3`,
`get_style/3`) accepts either A1-notation (`"B2"`) or a 1-indexed
`{row, col}` integer tuple (`{2, 2}`). Useful when porting from openpyxl
or iterating by numeric coordinates. `ExVEx.Utils.Coordinate.to_string/1`
is also public if you need to convert explicitly.

### Quality gates

- 125 ExUnit tests, all passing.
- `mix compile --warnings-as-errors`, `mix format --check-formatted`,
  `mix credo --strict` — all clean.
- GitHub Actions CI runs the above plus dialyzer on every push / PR.
- Output produced by `put_cell/4` + `save/2` is successfully read back by
  [umya-spreadsheet](https://crates.io/crates/umya-spreadsheet) (Rust) — a
  strong proxy for "Excel accepts this" without requiring Excel in CI.
