defmodule ExVExTest do
  use ExUnit.Case, async: true

  alias ExVEx.Packaging.{ContentTypes, Zip}
  alias ExVEx.Test.Fixtures
  alias ExVEx.Workbook

  describe "open/1" do
    test "returns a Workbook holding the package manifest and raw parts" do
      assert {:ok, %Workbook{} = book} = ExVEx.open(Fixtures.path("empty.xlsx"))

      assert %ContentTypes{} = book.content_types

      assert Enum.any?(book.content_types.overrides, fn o ->
               o.part_name == "/xl/workbook.xml"
             end)

      assert Map.has_key?(book.parts, "xl/workbook.xml")
      assert Map.has_key?(book.parts, "xl/worksheets/sheet1.xml")
    end

    test "returns an error for a non-existent file" do
      assert {:error, _} = ExVEx.open(Fixtures.tmp_path("nope.xlsx"))
    end
  end

  describe "sheet_names/1" do
    test "lists sheets in workbook.xml declaration order" do
      {:ok, book} = ExVEx.open(Fixtures.path("empty.xlsx"))
      assert ExVEx.sheet_names(book) == ["Sheet1", "Sheet2", "Sheet3"]
    end

    test "works on a workbook with actual cell content" do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      assert ExVEx.sheet_names(book) == ["Sheet1", "Sheet2", "Sheet3"]
    end
  end

  describe "sheet_path/2" do
    test "returns the package path of a sheet by name" do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))

      assert {:ok, "xl/worksheets/sheet1.xml"} = ExVEx.sheet_path(book, "Sheet1")
      assert {:ok, "xl/worksheets/sheet3.xml"} = ExVEx.sheet_path(book, "Sheet3")
    end

    test "returns :error for an unknown sheet name" do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      assert :error = ExVEx.sheet_path(book, "NoSuchSheet")
    end
  end

  describe "get_cell/3" do
    setup do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      %{book: book}
    end

    test "resolves shared-string cells to their text", %{book: book} do
      assert ExVEx.get_cell(book, "Sheet1", "A1") == {:ok, "A1"}
      assert ExVEx.get_cell(book, "Sheet1", "A4") == {:ok, "A4"}
      assert ExVEx.get_cell(book, "Sheet1", "B1") == {:ok, "B1"}
      assert ExVEx.get_cell(book, "Sheet1", "C3") == {:ok, "C3"}
    end

    test "preserves significant whitespace (xml:space=preserve)", %{book: book} do
      {:ok, value} = ExVEx.get_cell(book, "Sheet1", "C1")
      assert value =~ ~r/^\s*\nC1$/, "expected C1 to retain its literal leading newline"
    end

    test "returns {:ok, nil} for an empty cell", %{book: book} do
      assert ExVEx.get_cell(book, "Sheet1", "Z99") == {:ok, nil}
    end

    test "returns :error for an unknown sheet", %{book: book} do
      assert ExVEx.get_cell(book, "Ghost", "A1") == {:error, :unknown_sheet}
    end

    test "returns :error for an invalid coordinate", %{book: book} do
      assert ExVEx.get_cell(book, "Sheet1", "not-a-ref") == {:error, :invalid_coordinate}
    end
  end

  describe "get_cell/3 — numeric and boolean fixtures" do
    test "reads numeric and boolean cells from a synthetic fixture" do
      path =
        build_fixture(
          "types.xlsx",
          "cells.xlsx",
          fn parts ->
            sheet = Map.fetch!(parts, "xl/worksheets/sheet2.xml")
            new_sheet = inject_sheet_data(sheet, mixed_row_xml())
            Map.put(parts, "xl/worksheets/sheet2.xml", new_sheet)
          end
        )

      {:ok, book} = ExVEx.open(path)

      assert ExVEx.get_cell(book, "Sheet2", "A1") == {:ok, 42}
      assert ExVEx.get_cell(book, "Sheet2", "B1") == {:ok, 3.14}
      assert ExVEx.get_cell(book, "Sheet2", "C1") == {:ok, true}
      assert ExVEx.get_cell(book, "Sheet2", "D1") == {:ok, false}
      assert ExVEx.get_cell(book, "Sheet2", "E1") == {:ok, "pure inline"}
      assert ExVEx.get_cell(book, "Sheet2", "F1") == {:error, {:cell_error, "#REF!"}}
    end
  end

  defp build_fixture(name, source, transform) do
    out = Fixtures.tmp_path(name)
    on_exit(fn -> File.rm(out) end)

    {:ok, entries} = Zip.read(Fixtures.path(source))
    parts = Map.new(entries, fn e -> {e.path, e.data} end)
    parts = transform.(parts)

    mutated =
      Enum.map(entries, fn e ->
        %{e | data: Map.fetch!(parts, e.path)}
      end)

    :ok = Zip.write(out, mutated)
    out
  end

  defp inject_sheet_data(sheet_xml, row_xml) do
    String.replace(
      sheet_xml,
      "<sheetData/>",
      "<sheetData>#{row_xml}</sheetData>",
      global: false
    )
  end

  describe "put_cell/4 + save/2 round-trip" do
    setup do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      out = Fixtures.tmp_path("put_cell.xlsx")
      on_exit(fn -> File.rm(out) end)
      %{book: book, out: out}
    end

    test "overwrites a shared-string cell with a new inline string", %{book: book, out: out} do
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "Hello, ExVEx")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "A1") == {:ok, "Hello, ExVEx"}
      assert ExVEx.get_cell(reopened, "Sheet1", "A2") == {:ok, "A2"}
    end

    test "writes a number into a previously empty cell", %{book: book, out: out} do
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "E10", 42)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "F10", 3.14)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "E10") == {:ok, 42}
      assert ExVEx.get_cell(reopened, "Sheet1", "F10") == {:ok, 3.14}
    end

    test "writes booleans", %{book: book, out: out} do
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "G1", true)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "G2", false)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "G1") == {:ok, true}
      assert ExVEx.get_cell(reopened, "Sheet1", "G2") == {:ok, false}
    end

    test "clears a cell when given nil", %{book: book, out: out} do
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", nil)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "A1") == {:ok, nil}
      assert ExVEx.get_cell(reopened, "Sheet1", "A2") == {:ok, "A2"}
    end

    test "untouched sheets remain byte-identical to the source", %{book: book, out: out} do
      source_sheet2 = book.parts["xl/worksheets/sheet2.xml"]

      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "changed")
      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert reopened.parts["xl/worksheets/sheet2.xml"] == source_sheet2
      assert reopened.parts["xl/worksheets/sheet3.xml"] == book.parts["xl/worksheets/sheet3.xml"]
    end

    test "returns :error for an unknown sheet", %{book: book} do
      assert {:error, :unknown_sheet} = ExVEx.put_cell(book, "Ghost", "A1", "x")
    end

    test "returns :error for an invalid coordinate", %{book: book} do
      assert {:error, :invalid_coordinate} = ExVEx.put_cell(book, "Sheet1", "oops", "x")
    end
  end

  describe "formulas" do
    setup do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      out = Fixtures.tmp_path("formula.xlsx")
      on_exit(fn -> File.rm(out) end)
      %{book: book, out: out}
    end

    test "put_cell with {:formula, ...} writes a formula cell", %{book: book, out: out} do
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D1", 10)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D2", 20)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D3", {:formula, "=SUM(D1:D2)"})

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, "=SUM(D1:D2)"} = ExVEx.get_formula(reopened, "Sheet1", "D3")
    end

    test "put_cell with {:formula, ..., cached_value}", %{book: book, out: out} do
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D3", {:formula, "=1+1", 2})

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, 2} = ExVEx.get_cell(reopened, "Sheet1", "D3")
      assert {:ok, "=1+1"} = ExVEx.get_formula(reopened, "Sheet1", "D3")
    end

    test "get_formula returns nil for non-formula cells", %{book: book} do
      assert {:ok, nil} = ExVEx.get_formula(book, "Sheet1", "A1")
      assert {:ok, nil} = ExVEx.get_formula(book, "Sheet1", "Z99")
    end

    test "formulas in untouched cells on a mutated sheet survive", %{book: book, out: out} do
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A10", 100)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A11", 200)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A12", {:formula, "=SUM(A10:A11)", 300})

      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "Z99", "something unrelated")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, "=SUM(A10:A11)"} = ExVEx.get_formula(reopened, "Sheet1", "A12")
      assert {:ok, 300} = ExVEx.get_cell(reopened, "Sheet1", "A12")
      assert {:ok, "something unrelated"} = ExVEx.get_cell(reopened, "Sheet1", "Z99")
    end
  end

  describe "put_cell with Date / NaiveDateTime" do
    test "writes a Date and reads it back identically" do
      out = Fixtures.tmp_path("date_write.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D1", ~D[2024-01-15])
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D2", ~D[2000-02-29])

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "D1") == {:ok, ~D[2024-01-15]}
      assert ExVEx.get_cell(reopened, "Sheet1", "D2") == {:ok, ~D[2000-02-29]}
    end

    test "writes a NaiveDateTime and reads it back identically" do
      out = Fixtures.tmp_path("datetime_write.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D1", ~N[2024-01-15 12:00:00])

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, value} = ExVEx.get_cell(reopened, "Sheet1", "D1")
      assert NaiveDateTime.to_date(value) == ~D[2024-01-15]
      assert value.hour == 12
      assert value.minute == 0
    end

    test "untouched content still round-trips when date write mutates styles.xml" do
      out = Fixtures.tmp_path("date_preserves.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))

      source_sheet2 = book.parts["xl/worksheets/sheet2.xml"]

      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D1", ~D[2024-06-01])
      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert reopened.parts["xl/worksheets/sheet2.xml"] == source_sheet2
      assert ExVEx.get_cell(reopened, "Sheet1", "A1") == {:ok, "A1"}
    end
  end

  describe "put_cell with strings uses the shared-string table" do
    test "a repeated string value produces only one entry in sharedStrings.xml" do
      out = Fixtures.tmp_path("sst_dedup.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D1", "hello world")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D2", "hello world")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D3", "hello world")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "D1") == {:ok, "hello world"}
      assert ExVEx.get_cell(reopened, "Sheet1", "D2") == {:ok, "hello world"}
      assert ExVEx.get_cell(reopened, "Sheet1", "D3") == {:ok, "hello world"}

      sst_xml = reopened.parts["xl/sharedStrings.xml"]

      occurrences =
        sst_xml
        |> String.split("<si>")
        |> Enum.count(&String.contains?(&1, "hello world"))

      assert occurrences == 1,
             "expected 'hello world' to appear exactly once in sharedStrings.xml"
    end

    test "reuses an existing shared string entry when the text already appears" do
      out = Fixtures.tmp_path("sst_reuse.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))

      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "Z1", "A1")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "Z1") == {:ok, "A1"}

      sst_xml = reopened.parts["xl/sharedStrings.xml"]

      a1_count =
        sst_xml
        |> String.split("<si>")
        |> Enum.count(&String.match?(&1, ~r/<t[^>]*>A1<\/t>/))

      assert a1_count == 1, "'A1' already existed in the SST and should not be duplicated"
    end

    test "falls back to inline strings when the workbook has no SST" do
      out = Fixtures.tmp_path("no_sst.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("empty.xlsx"))
      refute book.shared_strings

      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "inline me")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert ExVEx.get_cell(reopened, "Sheet1", "A1") == {:ok, "inline me"}

      sheet_xml = reopened.parts["xl/worksheets/sheet1.xml"]
      assert sheet_xml =~ ~s(t="inlineStr")
    end
  end

  describe "get_style/3" do
    setup do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      %{book: book}
    end

    test "returns a default style for an unstyled cell", %{book: book} do
      assert {:ok, %ExVEx.Style{} = style} = ExVEx.get_style(book, "Sheet1", "A1")
      assert style.font.bold == false
      assert style.font.italic == false
      assert style.number_format == "General"
    end

    test "extracts font colour from a styled cell", %{book: book} do
      assert {:ok, %ExVEx.Style{font: font}} = ExVEx.get_style(book, "Sheet1", "B2")
      assert font.color.kind == :rgb
      assert font.color.value == "FFFF0000"
    end

    test "extracts border info from a bordered cell", %{book: book} do
      assert {:ok, %ExVEx.Style{border: border}} = ExVEx.get_style(book, "Sheet1", "C3")
      assert border.top.style == :thin
      assert border.bottom.style == :thin
      assert border.left.style == :thin
      assert border.right.style == :thin
    end

    test "extracts wrap-text alignment", %{book: book} do
      assert {:ok, %ExVEx.Style{alignment: alignment}} = ExVEx.get_style(book, "Sheet1", "C1")
      assert alignment.wrap_text == true
    end

    test "returns default style for empty cells", %{book: book} do
      assert {:ok, %ExVEx.Style{}} = ExVEx.get_style(book, "Sheet1", "Z99")
    end

    test "returns :error for an unknown sheet", %{book: book} do
      assert {:error, :unknown_sheet} = ExVEx.get_style(book, "Ghost", "A1")
    end
  end

  describe "merge_cells / unmerge_cells / merged_ranges" do
    setup do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      out = Fixtures.tmp_path("merge.xlsx")
      on_exit(fn -> File.rm(out) end)
      %{book: book, out: out}
    end

    test "merge_cells adds a range, merged_ranges lists it", %{book: book, out: out} do
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, ["A1:B2"]} = ExVEx.merged_ranges(reopened, "Sheet1")
    end

    test "default merge clears non-anchor cells but keeps the anchor", %{book: book, out: out} do
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, "A1"} = ExVEx.get_cell(reopened, "Sheet1", "A1")
      assert {:ok, nil} = ExVEx.get_cell(reopened, "Sheet1", "A2")
      assert {:ok, nil} = ExVEx.get_cell(reopened, "Sheet1", "B1")
      assert {:ok, nil} = ExVEx.get_cell(reopened, "Sheet1", "B2")

      assert {:ok, "A3"} = ExVEx.get_cell(reopened, "Sheet1", "A3")
    end

    test "preserve_values: true leaves every cell's value intact", %{book: book, out: out} do
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2", preserve_values: true)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, "A1"} = ExVEx.get_cell(reopened, "Sheet1", "A1")
      assert {:ok, "A2"} = ExVEx.get_cell(reopened, "Sheet1", "A2")
      assert {:ok, "B1"} = ExVEx.get_cell(reopened, "Sheet1", "B1")
      assert {:ok, "B2"} = ExVEx.get_cell(reopened, "Sheet1", "B2")

      assert {:ok, ["A1:B2"]} = ExVEx.merged_ranges(reopened, "Sheet1")
    end

    test "on_overlap: :error refuses an overlapping merge", %{book: book} do
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")

      assert {:error, {:overlaps, "A1:B2"}} =
               ExVEx.merge_cells(book, "Sheet1", "B2:C3")
    end

    test "on_overlap: :replace drops the overlapping range first", %{book: book, out: out} do
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "B2:C3", on_overlap: :replace)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, ["B2:C3"]} = ExVEx.merged_ranges(reopened, "Sheet1")
    end

    test "on_overlap: :allow lets partial overlaps coexist", %{book: book, out: out} do
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "B2:C3", on_overlap: :allow)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      {:ok, ranges} = ExVEx.merged_ranges(reopened, "Sheet1")
      assert Enum.sort(ranges) == ["A1:B2", "B2:C3"]
    end

    test "unmerge_cells removes an existing range", %{book: book, out: out} do
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")
      {:ok, book} = ExVEx.unmerge_cells(book, "Sheet1", "A1:B2")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, []} = ExVEx.merged_ranges(reopened, "Sheet1")
    end

    test "unmerge_cells errors on a range that isn't merged", %{book: book} do
      assert {:error, :not_merged} = ExVEx.unmerge_cells(book, "Sheet1", "A1:B2")
    end

    test "unmerge_cells with on_missing: :ignore is a no-op", %{book: book} do
      assert {:ok, ^book} = ExVEx.unmerge_cells(book, "Sheet1", "A1:B2", on_missing: :ignore)
    end

    test "merged_ranges on a sheet with no merges returns []", %{book: book} do
      assert {:ok, []} = ExVEx.merged_ranges(book, "Sheet1")
    end

    test "unmerged non-anchor cells stay empty (Excel convention; values not restored)", %{
      book: book,
      out: out
    } do
      {:ok, book} = ExVEx.merge_cells(book, "Sheet1", "A1:B2")
      {:ok, book} = ExVEx.unmerge_cells(book, "Sheet1", "A1:B2")

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, "A1"} = ExVEx.get_cell(reopened, "Sheet1", "A1")
      assert {:ok, nil} = ExVEx.get_cell(reopened, "Sheet1", "A2")
    end

    test "merging an invalid range returns :invalid_range", %{book: book} do
      assert {:error, :invalid_range} = ExVEx.merge_cells(book, "Sheet1", "not a range")
    end

    test "merging on an unknown sheet returns :unknown_sheet", %{book: book} do
      assert {:error, :unknown_sheet} = ExVEx.merge_cells(book, "Ghost", "A1:B2")
    end
  end

  describe "multi-sheet writes" do
    test "put_cell can target multiple sheets in sequence and all survive save" do
      out = Fixtures.tmp_path("multi_sheet.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))

      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "value on Sheet1")
      {:ok, book} = ExVEx.put_cell(book, "Sheet2", "B2", 42)
      {:ok, book} = ExVEx.put_cell(book, "Sheet3", "C3", {:formula, "=Sheet1!A1"})
      {:ok, book} = ExVEx.put_cell(book, "Sheet2", "B3", true)

      :ok = ExVEx.save(book, out)
      {:ok, reopened} = ExVEx.open(out)

      assert {:ok, "value on Sheet1"} = ExVEx.get_cell(reopened, "Sheet1", "A1")
      assert {:ok, 42} = ExVEx.get_cell(reopened, "Sheet2", "B2")
      assert {:ok, true} = ExVEx.get_cell(reopened, "Sheet2", "B3")
      assert {:ok, "=Sheet1!A1"} = ExVEx.get_formula(reopened, "Sheet3", "C3")
    end
  end

  describe "cells/2 + each_cell/2" do
    setup do
      {:ok, book} = ExVEx.open(Fixtures.path("cells.xlsx"))
      %{book: book}
    end

    test "cells/2 returns a map of every populated cell in the sheet", %{book: book} do
      assert {:ok, cells} = ExVEx.cells(book, "Sheet1")
      assert is_map(cells)

      assert cells["A1"] == "A1"
      assert cells["A4"] == "A4"
      assert cells["B1"] == "B1"
      refute Map.has_key?(cells, "Z99"), "empty cells must not appear"
    end

    test "each_cell/2 streams {ref, value} pairs in row-major order", %{book: book} do
      {:ok, stream} = ExVEx.each_cell(book, "Sheet1")
      list = Enum.to_list(stream)

      # every element is a {ref, value} pair
      assert Enum.all?(list, fn
               {ref, _} when is_binary(ref) -> true
               _ -> false
             end)

      refs = Enum.map(list, fn {ref, _} -> ref end)

      # row 1 comes before row 2, and A before B within a row
      assert Enum.find_index(refs, &(&1 == "A1")) < Enum.find_index(refs, &(&1 == "B1"))
      assert Enum.find_index(refs, &(&1 == "A1")) < Enum.find_index(refs, &(&1 == "A2"))
    end

    test "cells/2 on an unknown sheet returns :error", %{book: book} do
      assert {:error, :unknown_sheet} = ExVEx.cells(book, "Ghost")
    end
  end

  describe "get_cell/3 — date detection" do
    test "numeric cell with a built-in date numFmtId returns a Date" do
      path =
        build_fixture("dates.xlsx", "cells.xlsx", fn parts ->
          parts
          |> Map.put("xl/styles.xml", date_styles_xml())
          |> Map.put("xl/worksheets/sheet1.xml", date_sheet_xml())
        end)

      {:ok, book} = ExVEx.open(path)

      assert ExVEx.get_cell(book, "Sheet1", "A1") == {:ok, ~D[2024-01-15]}
      assert ExVEx.get_cell(book, "Sheet1", "A2") == {:ok, ~D[1900-01-02]}
      assert ExVEx.get_cell(book, "Sheet1", "B1") == {:ok, 42}
    end

    test "numeric cell with a custom date format returns a Date" do
      path =
        build_fixture("custom_dates.xlsx", "cells.xlsx", fn parts ->
          parts
          |> Map.put("xl/styles.xml", custom_date_styles_xml())
          |> Map.put("xl/worksheets/sheet1.xml", date_sheet_xml())
        end)

      {:ok, book} = ExVEx.open(path)

      assert ExVEx.get_cell(book, "Sheet1", "A1") == {:ok, ~D[2024-01-15]}
    end

    test "numeric cell with a date+time format returns a NaiveDateTime" do
      path =
        build_fixture("datetimes.xlsx", "cells.xlsx", fn parts ->
          parts
          |> Map.put("xl/styles.xml", datetime_styles_xml())
          |> Map.put("xl/worksheets/sheet1.xml", datetime_sheet_xml())
        end)

      {:ok, book} = ExVEx.open(path)

      {:ok, value} = ExVEx.get_cell(book, "Sheet1", "A1")
      assert %NaiveDateTime{} = value
      assert NaiveDateTime.to_date(value) == ~D[2024-01-15]
      assert value.hour == 12
      assert value.minute == 0
    end
  end

  defp date_styles_xml do
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      </cellXfs>
    </styleSheet>
    """
  end

  defp custom_date_styles_xml do
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      </cellXfs>
    </styleSheet>
    """
  end

  defp datetime_styles_xml do
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd h:mm"/></numFmts>
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
      </cellXfs>
    </styleSheet>
    """
  end

  defp date_sheet_xml do
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" s="1"><v>45306</v></c>
          <c r="B1"><v>42</v></c>
        </row>
        <row r="2">
          <c r="A2" s="1"><v>2</v></c>
        </row>
      </sheetData>
    </worksheet>
    """
  end

  defp datetime_sheet_xml do
    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetData>
        <row r="1">
          <c r="A1" s="1"><v>45306.5</v></c>
        </row>
      </sheetData>
    </worksheet>
    """
  end

  defp mixed_row_xml do
    """
    <row r="1">
      <c r="A1"><v>42</v></c>
      <c r="B1"><v>3.14</v></c>
      <c r="C1" t="b"><v>1</v></c>
      <c r="D1" t="b"><v>0</v></c>
      <c r="E1" t="inlineStr"><is><t>pure inline</t></is></c>
      <c r="F1" t="e"><v>#REF!</v></c>
    </row>
    """
  end

  describe "save/2 (round-trip identity on an untouched workbook)" do
    test "open -> save -> re-open yields the same parts" do
      out = Fixtures.tmp_path("roundtrip.xlsx")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("empty.xlsx"))
      :ok = ExVEx.save(book, out)

      {:ok, round_tripped} = ExVEx.open(out)

      assert round_tripped.parts == book.parts
    end

    test "preserves the vbaProject.bin macro blob byte-for-byte" do
      out = Fixtures.tmp_path("macros_roundtrip.xlsm")
      on_exit(fn -> File.rm(out) end)

      {:ok, book} = ExVEx.open(Fixtures.path("with_macros.xlsm"))
      :ok = ExVEx.save(book, out)

      {:ok, round_tripped} = ExVEx.open(out)

      vba_key =
        Enum.find(Map.keys(book.parts), fn path ->
          String.ends_with?(path, "vbaProject.bin")
        end)

      assert vba_key, "fixture should contain a vbaProject.bin entry"
      assert round_tripped.parts[vba_key] == book.parts[vba_key]
    end
  end
end
