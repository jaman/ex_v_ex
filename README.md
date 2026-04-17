# ExVEx

**Elixir vs. Excel.** A pure-Elixir library for reading and editing existing
`.xlsx` and `.xlsm` files with round-trip fidelity — no Rust, no Python, no
NIFs, no second runtime in your deployment.

```elixir
{:ok, book} = ExVEx.open("inventory.xlsx")

# Read
ExVEx.sheet_names(book)                           #=> ["Sheet1", "Sheet2"]
ExVEx.get_cell(book, "Sheet1", "A1")              #=> {:ok, "widget"}
ExVEx.get_cell(book, "Sheet1", "B2")              #=> {:ok, 42}
ExVEx.get_cell(book, "Sheet1", "C3")              #=> {:ok, ~D[2024-01-15]}
ExVEx.get_formula(book, "Sheet1", "D4")           #=> {:ok, "=SUM(B2:B10)"}
ExVEx.get_style(book, "Sheet1", "A1")             #=> {:ok, %ExVEx.Style{...}}

# Write
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D1", "Updated")
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D2", 3.14)
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D3", true)
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D4", ~D[2024-06-01])
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D5", {:formula, "=SUM(A1:A10)"})

# Coordinates accept both A1 refs and {row, col} integer tuples
ExVEx.get_cell(book, "Sheet1", {2, 3})            #=> same as ExVEx.get_cell(book, "Sheet1", "C2")
{:ok, book} = ExVEx.put_cell(book, "Sheet1", {10, 5}, "tuple write")

# Merge / unmerge
{:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")
{:ok, ["A1:B2"]} = ExVEx.merged_ranges(book, "Sheet1")
{:ok, book} = ExVEx.unmerge_cells(book, "Sheet1", "A1:B2")

# Save
:ok = ExVEx.save(book, "inventory.xlsx")
```

## Why

The Elixir ecosystem has `Elixlsx` (write-only) and `xlsx_reader` (read-only),
but no first-class story for *editing existing* spreadsheets. Every team that
needs this today reaches for Python (`openpyxl`) or a Rust NIF — which drags
a second runtime into the deployment.

ExVEx fills that gap in pure Elixir.

## Status

**v0.1 — pre-alpha.** Core read/write/round-trip is solid and externally
validated against [umya-spreadsheet](https://crates.io/crates/umya-spreadsheet)
(Rust). 119 tests, zero credo issues, zero compile warnings.

### What works

- Open `.xlsx` and `.xlsm` files
- Round-trip identity on untouched content (unknown XML, custom parts, VBA
  macros all pass through byte-for-byte)
- Read cell values: strings (shared + inline), numbers, booleans, dates,
  date-times, formula results, cell errors
- Read cell styles: font, fill, border, alignment, number format
- Write cell values: strings (through the SST with automatic dedup),
  numbers, booleans, `nil` (clear), `Date`, `NaiveDateTime`, formulas (with
  or without cached value)
- Sheet navigation: `sheet_names/1`, `sheet_path/2`
- Bulk reads: `cells/2` (map), `each_cell/2` (stream in row-major order)
- Get cell formula: `get_formula/3`
- Merge / unmerge cells with configurable overlap and value-preservation
  behaviour (`merge_cells/3,4`, `unmerge_cells/3,4`, `merged_ranges/2`)

### Not yet

- Style mutation (set bold, change font, etc.)
- Row/column insertion, defined names
- Charts, images, pivot tables, comments

## Installation

```elixir
def deps do
  [{:ex_v_ex, "~> 0.1.0"}]
end
```

## Design

Three commitments drive the design:

1. **Lazy raw parts.** The archive is read into a flat `%{path => bytes}`
   map. Parts you never touch are written back byte-identical on save.
2. **Immutable functional API.** Every mutating operation returns a new
   `%ExVEx.Workbook{}`. No GenServers, no mutable references.
3. **Preserve over re-serialize.** When a part is mutated (a worksheet with
   a written cell; the styles.xml when a date write adds an xf; the
   sharedStrings.xml when a new string is interned), ExVEx surgically edits
   just the changed sub-tree and leaves the surrounding XML alone —
   namespaces, unknown attributes, and extension elements survive.

This is the same strategy used by [umya-spreadsheet](https://crates.io/crates/umya-spreadsheet)
(Rust), and stronger than [edit-xlsx](https://crates.io/crates/edit-xlsx)'s
full-deserialize-reserialize approach. It's also the only way to preserve
`.xlsm` VBA binaries byte-for-byte, which ExVEx verifies in its test suite.

## License

MIT. See [LICENSE](LICENSE).
