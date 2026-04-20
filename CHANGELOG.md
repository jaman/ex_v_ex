# Changelog

All notable changes to this project will be documented here. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added (workbook creation)

ExVEx is now a complete xlsx library — it can create workbooks from scratch,
not just edit existing ones. No Python or Rust dependency needed for
template generation.

- `ExVEx.new/0` — returns a minimal blank workbook with a single empty
  sheet named `"Sheet1"`. No fixture file needed; the skeleton parts are
  embedded in the library.
- `ExVEx.add_sheet(book, name)` — appends a new empty sheet.
  `{:error, :duplicate_sheet_name}` if the name is already in use.
- `ExVEx.rename_sheet(book, old, new)` — changes a sheet's name in place.
  `{:error, :unknown_sheet}` or `{:error, :duplicate_sheet_name}` as
  appropriate. Same-name is a no-op.
- `ExVEx.remove_sheet(book, name)` — drops a sheet and its worksheet part,
  plus the matching Content Types Override and workbook relationship.
  `{:error, :last_sheet}` guards against producing an invalid
  zero-sheet workbook.

All four coordinate `workbook.xml`, `workbook.xml.rels`, and
`[Content_Types].xml` together on every change. A new
`OOXML.Workbook.serialize_into/2` rewrites the `<sheets>` section while
preserving surrounding elements at the SimpleForm level.

Example — build a multi-sheet template from zero:

    {:ok, book} = ExVEx.new()
    {:ok, book} = ExVEx.rename_sheet(book, "Sheet1", "Summary")
    {:ok, book} = ExVEx.add_sheet(book, "Data")
    {:ok, book} = ExVEx.add_sheet(book, "Formulas")
    {:ok, book} = ExVEx.put_cell(book, "Data", "A1", "alpha")
    {:ok, book} = ExVEx.put_cell(book, "Formulas", "A1", {:formula, "=SUM(Data!A1:A10)"})
    :ok = ExVEx.save(book, "template.xlsx")

### Changed (performance)

Bulk writes are now ~20× faster with ~60× less memory churn.

- **Sheet tree cache.** `%ExVEx.Workbook{}` now caches the parsed
  `Saxy.SimpleForm` tree for every worksheet on first access. Subsequent
  `get_cell`, `put_cell`, `merge_cells`, `unmerge_cells`, `merged_ranges`,
  `get_formula`, `get_style`, `cells`, and `each_cell` calls reuse the
  cached tree instead of re-parsing the XML. `save/2` re-serializes only
  dirty sheet trees once at flush time.
- **Shared-string interns are O(1).** `ExVEx.OOXML.SharedStrings` now
  stores strings in two maps (`by_index`, `by_string`) instead of a
  tuple. Interning a new string was O(N) per call (tuple copy); it is
  now O(1).

Benchmark: 500 `put_cell + save` dropped from 142 ms / 348 MB to 7 ms /
6 MB. 1000 unique string interns + save dropped from 761 ms / 1.8 GB to
36 ms / 16 MB. See `bench/results/README.md` for the full report and
instructions on reproducing.

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
