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

# Write
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D1", "Updated")
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D2", 3.14)
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D3", true)
{:ok, book} = ExVEx.put_cell(book, "Sheet1", "D4", {:formula, "=SUM(A1:A10)"})

# Save
:ok = ExVEx.save(book, "inventory.xlsx")
```

## Why

The Elixir ecosystem has `Elixlsx` (write-only) and `xlsx_reader` (read-only),
but no first-class story for *editing existing* spreadsheets. Every team that
needs this today reaches for Python (`openpyxl`) or a Rust NIF — which drags
a second runtime into the deployment. For one team, adding Python inflated a
50 MB release to 900 MB.

ExVEx fills that gap in pure Elixir.

## Status

**v0.1 — pre-alpha.** Core read/write/round-trip is solid and externally
validated against [umya-spreadsheet](https://crates.io/crates/umya-spreadsheet)
(Rust). See [CHANGELOG.md](CHANGELOG.md) for what works today and what's
next.

### What works

- Open `.xlsx` and `.xlsm` files
- Round-trip identity on untouched content (unknown XML, custom parts, VBA
  macros all pass through byte-for-byte)
- Read cell values: strings (shared + inline), numbers, booleans, dates,
  date-times, formula results, cell errors
- Write cell values: strings, numbers, booleans, `nil` (clear), formulas
  (with or without cached value)
- Sheet navigation: `sheet_names/1`, `sheet_path/2`
- Bulk reads: `cells/2` (map), `each_cell/2` (stream in row-major order)
- Get cell formula: `get_formula/3`

### Not yet

- Writing `Date` / `NaiveDateTime` values directly (requires styles.xml mutation)
- Writing through the shared strings table (current writes use inline strings)
- Cell style reads (font, fill, border, alignment)
- Style mutation
- Row/column insertion, merged cells, defined names, charts, images, pivot
  tables, comments

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
   a written cell), ExVEx surgically edits just the changed sub-tree and
   leaves the surrounding XML alone — namespaces, unknown attributes, and
   extension elements survive.

This is the same strategy used by [umya-spreadsheet](https://crates.io/crates/umya-spreadsheet)
(Rust), and stronger than [edit-xlsx](https://crates.io/crates/edit-xlsx)'s
full-deserialize-reserialize approach. It's also the only way to preserve
`.xlsm` VBA binaries byte-for-byte, which ExVEx verifies in its test suite.

## License

MIT. See [LICENSE](LICENSE).
