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
- OOXML parsers: `Packaging.ContentTypes`, `Packaging.Relationships`,
  `OOXML.Workbook`, `OOXML.Worksheet`, `OOXML.SharedStrings`, `OOXML.Styles`.
- Style model: `ExVEx.Style` + `Font`, `Fill`, `Border`, `Side`,
  `Alignment`, `Color`.
- Coordinate utilities: `ExVEx.Utils.Coordinate` — A1 ↔ `{row, col}`,
  Excel's bijective base-26 column labels.

### Quality gates

- 90 ExUnit tests, all passing.
- `mix compile --warnings-as-errors`, `mix format --check-formatted`,
  `mix credo --strict` — all clean.
- GitHub Actions CI runs the above plus dialyzer on every push / PR.
- Output produced by `put_cell/4` + `save/2` is successfully read back by
  [umya-spreadsheet](https://crates.io/crates/umya-spreadsheet) (Rust) — a
  strong proxy for "Excel accepts this" without requiring Excel in CI.
