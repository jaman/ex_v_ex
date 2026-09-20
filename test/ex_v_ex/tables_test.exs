defmodule ExVEx.TablesTest do
  use ExUnit.Case, async: true

  alias ExVEx.Table
  alias ExVEx.Test.Fixtures

  setup do
    out = Fixtures.tmp_path("tables.xlsx")
    on_exit(fn -> File.rm(out) end)
    %{out: out}
  end

  defp sales_book do
    {:ok, book} = ExVEx.new()

    rows = [
      ["Region", "Qty", "Amount"],
      ["North", 2, 10.5],
      ["South", 3, 20.0],
      ["East", 1, 5.25]
    ]

    Enum.reduce(Enum.with_index(rows, 1), book, fn {row, r}, acc ->
      Enum.reduce(Enum.with_index(row, 1), acc, fn {value, c}, acc2 ->
        {:ok, acc2} = ExVEx.put_cell(acc2, "Sheet1", {r, c}, value)
        acc2
      end)
    end)
  end

  defp round_trip(book, out) do
    :ok = ExVEx.save(book, out)
    {:ok, reopened} = ExVEx.open(out)
    reopened
  end

  defp part(book, path), do: Map.fetch!(book.parts, path)

  describe "add_table/4" do
    test "creates a table over existing headers and links every package piece", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      reopened = round_trip(book, out)

      assert {:ok, %Table{} = table} = ExVEx.table(reopened, "Sales")
      assert table.sheet == "Sheet1"
      assert table.ref == "A1:C4"
      assert table.header_range == "A1:C1"
      assert table.data_range == "A2:C4"
      assert table.totals_range == nil
      assert table.columns == ["Region", "Qty", "Amount"]
      assert table.style == "TableStyleMedium2"
      assert table.show_row_stripes == true

      assert part(reopened, "xl/tables/table1.xml") =~
               ~s(name="Sales" displayName="Sales" ref="A1:C4")

      assert part(reopened, "xl/worksheets/_rels/sheet1.xml.rels") =~
               ~s(Target="../tables/table1.xml")

      assert part(reopened, "xl/worksheets/sheet1.xml") =~
               ~s(<tableParts count="1"><tablePart r:id="rId1"/></tableParts>)

      assert part(reopened, "xl/worksheets/sheet1.xml") =~
               ~s(xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships")

      assert part(reopened, "[Content_Types].xml") =~
               ~s(PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
    end

    test "explicit column names are written into the header cells", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B2", 1)

      {:ok, book} =
        ExVEx.add_table(book, "Sheet1", "B1:C3",
          name: "T",
          columns: ["Id", "Value"],
          style: "TableStyleLight1"
        )

      reopened = round_trip(book, out)
      assert ExVEx.get_cell(reopened, "Sheet1", "B1") == {:ok, "Id"}
      assert ExVEx.get_cell(reopened, "Sheet1", "C1") == {:ok, "Value"}

      assert {:ok, %Table{columns: ["Id", "Value"], style: "TableStyleLight1"}} =
               ExVEx.table(reopened, "T")
    end

    test "empty header cells get placeholder names and non-text cells become text", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "Name")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B1", 42)
      {:ok, book} = ExVEx.add_table(book, "Sheet1", "A1:C2", name: "T")

      reopened = round_trip(book, out)
      assert {:ok, %Table{columns: ["Name", "42", "Column3"]}} = ExVEx.table(reopened, "T")
      assert ExVEx.get_cell(reopened, "Sheet1", "B1") == {:ok, "42"}
      assert ExVEx.get_cell(reopened, "Sheet1", "C1") == {:ok, "Column3"}
    end

    test "default name is TableN", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4")
      reopened = round_trip(book, out)
      assert [%Table{name: "Table1"}] = ExVEx.tables(reopened)
    end

    test "second table gets a fresh id, part, and relationship", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "First")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "E1", "K")
      {:ok, book} = ExVEx.add_table(book, "Sheet1", "E1:E3", name: "Second")
      reopened = round_trip(book, out)

      assert Enum.map(ExVEx.tables(reopened), & &1.name) == ["First", "Second"]
      assert part(reopened, "xl/tables/table2.xml") =~ ~s(id="2" name="Second")

      assert part(reopened, "xl/worksheets/sheet1.xml") =~
               ~s(<tablePart r:id="rId1"/><tablePart r:id="rId2"/>)
    end

    test "rejects invalid input" do
      book = sales_book()
      {:ok, book_with} = ExVEx.add_table(book, "Sheet1", "A1:C4", name: "Sales")
      {:ok, merged} = ExVEx.merge_cells(book, "Sheet1", "E1:F1")
      {:ok, named} = ExVEx.define_name(book, "Sales", "Sheet1!$A$1")

      assert ExVEx.add_table(book, "Nope", "A1:C4") == {:error, :unknown_sheet}
      assert ExVEx.add_table(book, "Sheet1", "junk") == {:error, :invalid_range}
      assert ExVEx.add_table(book, "Sheet1", "A1:C1") == {:error, :range_too_small}

      assert ExVEx.add_table(book, "Sheet1", "A1:C4", name: "Bad Name") ==
               {:error, {:invalid_table_name, "Bad Name"}}

      assert ExVEx.add_table(book, "Sheet1", "A1:C4", name: "AB12") ==
               {:error, {:invalid_table_name, "AB12"}}

      assert ExVEx.add_table(book_with, "Sheet1", "E1:F3", name: "sales") ==
               {:error, {:duplicate_table_name, "sales"}}

      assert ExVEx.add_table(named, "Sheet1", "A1:C4", name: "Sales") ==
               {:error, {:duplicate_table_name, "Sales"}}

      assert ExVEx.add_table(book_with, "Sheet1", "B2:D6", name: "Other") ==
               {:error, {:overlaps_table, "Sales"}}

      assert ExVEx.add_table(merged, "Sheet1", "E1:F3", name: "T") ==
               {:error, {:overlaps_merged_range, "E1:F1"}}

      assert ExVEx.add_table(book, "Sheet1", "A1:C4", columns: ["a", "b"]) ==
               {:error, :column_count_mismatch}

      assert ExVEx.add_table(book, "Sheet1", "A1:C4", columns: ["a", "A", "c"]) ==
               {:error, {:duplicate_column_name, "A"}}
    end
  end

  describe "tables/1 and tables/2" do
    test "list per workbook and per sheet" do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.add_sheet(book, "Other")
      {:ok, book} = ExVEx.put_cell(book, "Other", "A1", "H")
      {:ok, book} = ExVEx.add_table(book, "Other", "A1:A2", name: "Second")

      assert Enum.map(ExVEx.tables(book), &{&1.name, &1.sheet}) == [
               {"Sales", "Sheet1"},
               {"Second", "Other"}
             ]

      assert {:ok, [%Table{name: "Second"}]} = ExVEx.tables(book, "Other")
      assert ExVEx.tables(book, "Nope") == {:error, :unknown_sheet}
      assert ExVEx.table(book, "nothing") == {:error, :unknown_table}
    end
  end

  describe "remove_table/2" do
    test "drops the part, relationship, link, and content type; cells stay", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.remove_table(book, "Sales")
      reopened = round_trip(book, out)

      assert ExVEx.tables(reopened) == []
      refute Map.has_key?(reopened.parts, "xl/tables/table1.xml")
      refute Map.has_key?(reopened.parts, "xl/worksheets/_rels/sheet1.xml.rels")
      refute part(reopened, "xl/worksheets/sheet1.xml") =~ "tableParts"
      refute part(reopened, "[Content_Types].xml") =~ "table1.xml"
      assert ExVEx.get_cell(reopened, "Sheet1", "C3") == {:ok, 20.0}
      assert ExVEx.remove_table(reopened, "Sales") == {:error, :unknown_table}
    end

    test "structured references to the table become plain ranges", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "E1", {:formula, "=SUM(Sales[Amount])"})
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "E2", {:formula, "=Sales[[#Headers],[Qty]]"})
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "C3", {:formula, "=[@Qty]*2"})
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "E4", {:formula, "=COUNTA(Sales[#All])"})
      {:ok, book} = ExVEx.add_sheet(book, "Other")
      {:ok, book} = ExVEx.put_cell(book, "Other", "A1", {:formula, "=SUM(Sales[[Qty]:[Amount]])"})
      {:ok, book} = ExVEx.remove_table(book, "Sales")
      reopened = round_trip(book, out)

      assert ExVEx.get_formula(reopened, "Sheet1", "E1") == {:ok, "=SUM(C2:C4)"}
      assert ExVEx.get_formula(reopened, "Sheet1", "E2") == {:ok, "=B1"}
      assert ExVEx.get_formula(reopened, "Sheet1", "C3") == {:ok, "=B3*2"}
      assert ExVEx.get_formula(reopened, "Sheet1", "E4") == {:ok, "=COUNTA(A1:C4)"}
      assert ExVEx.get_formula(reopened, "Other", "A1") == {:ok, "=SUM(Sheet1!B2:C4)"}
    end
  end

  describe "rename_table/3" do
    test "renames the part and every qualified reference", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "E1", {:formula, "=SUM(Sales[Amount])"})
      {:ok, book} = ExVEx.define_name(book, "Total", "SUM(Sales[Amount])")
      {:ok, book} = ExVEx.rename_table(book, "Sales", "Revenue")
      reopened = round_trip(book, out)

      assert {:ok, %Table{name: "Revenue"}} = ExVEx.table(reopened, "Revenue")
      assert ExVEx.get_formula(reopened, "Sheet1", "E1") == {:ok, "=SUM(Revenue[Amount])"}
      assert [%{name: "Total", reference: "SUM(Revenue[Amount])"}] = ExVEx.defined_names(reopened)

      assert ExVEx.rename_table(reopened, "Revenue", "bad name") ==
               {:error, {:invalid_table_name, "bad name"}}
    end
  end

  describe "rename_table_column/4" do
    test "updates the header cell, the part, and formulas", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "E1", {:formula, "=SUM(Sales[Amount])"})
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D2", {:formula, "=[@Amount]*2"})
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "E5", {:formula, "=[@Amount]*2"})
      {:ok, book} = ExVEx.rename_table_column(book, "Sales", "Amount", "Total")
      reopened = round_trip(book, out)

      assert ExVEx.get_cell(reopened, "Sheet1", "C1") == {:ok, "Total"}
      assert {:ok, %Table{columns: ["Region", "Qty", "Total"]}} = ExVEx.table(reopened, "Sales")
      assert ExVEx.get_formula(reopened, "Sheet1", "E1") == {:ok, "=SUM(Sales[Total])"}
      assert ExVEx.get_formula(reopened, "Sheet1", "D2") == {:ok, "=[@Amount]*2"}
      assert ExVEx.get_formula(reopened, "Sheet1", "E5") == {:ok, "=[@Amount]*2"}

      assert ExVEx.rename_table_column(reopened, "Sales", "Nope", "X") ==
               {:error, {:unknown_column, "Nope"}}

      assert ExVEx.rename_table_column(reopened, "Sales", "Qty", "total") ==
               {:error, {:duplicate_column_name, "total"}}
    end

    test "implicit references inside the table are renamed", %{out: out} do
      {:ok, book} = ExVEx.new()
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A1", "Qty")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B1", "Double")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "A2", 2)
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "B2", {:formula, "=[@Qty]*2"})
      {:ok, book} = ExVEx.add_table(book, "Sheet1", "A1:B2", name: "T")
      {:ok, book} = ExVEx.rename_table_column(book, "T", "Qty", "Count")
      reopened = round_trip(book, out)
      assert ExVEx.get_formula(reopened, "Sheet1", "B2") == {:ok, "=[@Count]*2"}
    end
  end

  describe "resize_table/3" do
    test "growing adds columns from header cells and rows; shrinking drops them", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_cell(book, "Sheet1", "D1", "Notes")
      {:ok, grown} = ExVEx.resize_table(book, "Sales", "A1:E6")
      reopened = round_trip(grown, out)

      assert {:ok, %Table{ref: "A1:E6", columns: ["Region", "Qty", "Amount", "Notes", "Column5"]}} =
               ExVEx.table(reopened, "Sales")

      assert ExVEx.get_cell(reopened, "Sheet1", "E1") == {:ok, "Column5"}

      {:ok, shrunk} = ExVEx.resize_table(reopened, "Sales", "A1:B3")

      assert {:ok, %Table{ref: "A1:B3", columns: ["Region", "Qty"]}} =
               ExVEx.table(shrunk, "Sales")

      assert ExVEx.resize_table(shrunk, "Sales", "A2:B3") == {:error, :header_row_must_stay}
      assert ExVEx.resize_table(shrunk, "Sales", "A1:B1") == {:error, :range_too_small}
    end
  end

  describe "put_table_style/3" do
    test "changes the style name and banding flags", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")

      {:ok, book} =
        ExVEx.put_table_style(book, "Sales",
          style: "TableStyleDark3",
          show_row_stripes: false,
          show_first_column: true
        )

      reopened = round_trip(book, out)

      assert {:ok,
              %Table{style: "TableStyleDark3", show_row_stripes: false, show_first_column: true}} =
               ExVEx.table(reopened, "Sales")
    end
  end

  describe "totals row" do
    test "put_table_totals_row/3 adds the row with SUBTOTAL formulas and labels", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")

      {:ok, book} =
        ExVEx.put_table_totals_row(book, "Sales",
          Region: {:label, "Total"},
          Qty: :count,
          Amount: :sum
        )

      reopened = round_trip(book, out)

      assert {:ok, %Table{ref: "A1:C5", data_range: "A2:C4", totals_range: "A5:C5"}} =
               ExVEx.table(reopened, "Sales")

      assert ExVEx.get_cell(reopened, "Sheet1", "A5") == {:ok, "Total"}
      assert ExVEx.get_formula(reopened, "Sheet1", "B5") == {:ok, "SUBTOTAL(103,Sales[Qty])"}
      assert ExVEx.get_formula(reopened, "Sheet1", "C5") == {:ok, "SUBTOTAL(109,Sales[Amount])"}
      assert part(reopened, "xl/tables/table1.xml") =~ ~s(totalsRowCount="1")

      assert part(reopened, "xl/tables/table1.xml") =~
               ~s(<tableColumn id="3" name="Amount" totalsRowFunction="sum"/>)

      assert part(reopened, "xl/tables/table1.xml") =~ ~s(<autoFilter ref="A1:C4"/>)
      refute part(reopened, "xl/tables/table1.xml") =~ ~s(totalsRowShown="0")
    end

    test "custom formulas and updates to an existing totals row", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_table_totals_row(book, "Sales", Amount: :sum)
      {:ok, book} = ExVEx.put_table_totals_row(book, "Sales", Qty: {:custom, "MAX(Sales[Qty])*2"})
      reopened = round_trip(book, out)

      assert {:ok, %Table{ref: "A1:C5"}} = ExVEx.table(reopened, "Sales")
      assert ExVEx.get_formula(reopened, "Sheet1", "B5") == {:ok, "MAX(Sales[Qty])*2"}
      assert ExVEx.get_formula(reopened, "Sheet1", "C5") == {:ok, "SUBTOTAL(109,Sales[Amount])"}

      assert part(reopened, "xl/tables/table1.xml") =~
               ~s|<totalsRowFormula>MAX(Sales[Qty])*2</totalsRowFormula>|
    end

    test "refuses when the row below is occupied and rejects unknown columns" do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, occupied} = ExVEx.put_cell(book, "Sheet1", "B5", "x")

      assert ExVEx.put_table_totals_row(occupied, "Sales", Amount: :sum) ==
               {:error, {:occupied, "A5:C5"}}

      assert ExVEx.put_table_totals_row(book, "Sales", Nope: :sum) ==
               {:error, {:unknown_column, "Nope"}}
    end

    test "remove_table_totals_row/2 clears the cells and shrinks the ref", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_table_totals_row(book, "Sales", Amount: :sum)
      {:ok, book} = ExVEx.remove_table_totals_row(book, "Sales")
      reopened = round_trip(book, out)

      assert {:ok, %Table{ref: "A1:C4", totals_range: nil}} = ExVEx.table(reopened, "Sales")
      assert ExVEx.get_cell(reopened, "Sheet1", "C5") == {:ok, nil}
      refute part(reopened, "xl/tables/table1.xml") =~ "totalsRowFunction"
    end
  end

  describe "reading rows" do
    test "table_rows/2 and table_records/2 return the data body only", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_table_totals_row(book, "Sales", Amount: :sum)
      reopened = round_trip(book, out)

      assert ExVEx.table_rows(reopened, "Sales") ==
               {:ok, [["North", 2, 10.5], ["South", 3, 20.0], ["East", 1, 5.25]]}

      assert {:ok, [%{"Region" => "North", "Qty" => 2, "Amount" => 10.5} | _]} =
               ExVEx.table_records(reopened, "Sales")

      assert ExVEx.table_rows(reopened, "Nope") == {:error, :unknown_table}
    end
  end

  describe "append_table_rows/3" do
    test "extends the table below the data and moves the totals row down", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_table_totals_row(book, "Sales", Amount: :sum)

      {:ok, book} =
        ExVEx.append_table_rows(book, "Sales", [
          ["West", 4, 1.0],
          %{"Region" => "Central", "Amount" => 2.0}
        ])

      reopened = round_trip(book, out)

      assert {:ok, %Table{ref: "A1:C7", data_range: "A2:C6", totals_range: "A7:C7"}} =
               ExVEx.table(reopened, "Sales")

      assert ExVEx.get_cell(reopened, "Sheet1", "A5") == {:ok, "West"}
      assert ExVEx.get_cell(reopened, "Sheet1", "A6") == {:ok, "Central"}
      assert ExVEx.get_cell(reopened, "Sheet1", "B6") == {:ok, nil}
      assert ExVEx.get_formula(reopened, "Sheet1", "C7") == {:ok, "SUBTOTAL(109,Sales[Amount])"}
      assert ExVEx.get_formula(reopened, "Sheet1", "C5") == {:ok, nil}
      assert part(reopened, "xl/tables/table1.xml") =~ ~s(ref="A1:C7")
      assert part(reopened, "xl/tables/table1.xml") =~ ~s(<autoFilter ref="A1:C6"/>)
    end

    test "refuses to overwrite occupied cells and unknown columns" do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, occupied} = ExVEx.put_cell(book, "Sheet1", "C6", "x")

      assert ExVEx.append_table_rows(occupied, "Sales", [[1, 2, 3], [4, 5, 6]]) ==
               {:error, {:occupied, "A5:C6"}}

      assert ExVEx.append_table_rows(book, "Sales", [%{"Nope" => 1}]) ==
               {:error, {:unknown_column, "Nope"}}

      assert ExVEx.append_table_rows(book, "Sales", [[1, 2, 3, 4]]) ==
               {:error, :column_count_mismatch}

      assert ExVEx.append_table_rows(book, "Sales", []) == {:ok, book}
    end
  end

  describe "structural shifts" do
    test "insert_row inside the table grows it", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.insert_row(book, "Sheet1", 3, 2)
      reopened = round_trip(book, out)
      assert {:ok, %Table{ref: "A1:C6"}} = ExVEx.table(reopened, "Sales")
    end

    test "insert_column inside adds a column and writes its header", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.insert_column(book, "Sheet1", 2)
      reopened = round_trip(book, out)

      assert {:ok, %Table{ref: "A1:D4", columns: ["Region", "Column2", "Qty", "Amount"]}} =
               ExVEx.table(reopened, "Sales")

      assert ExVEx.get_cell(reopened, "Sheet1", "B1") == {:ok, "Column2"}
      assert ExVEx.get_cell(reopened, "Sheet1", "C1") == {:ok, "Qty"}
    end

    test "delete_column inside drops the column", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.delete_column(book, "Sheet1", 2)
      reopened = round_trip(book, out)

      assert {:ok, %Table{ref: "A1:B4", columns: ["Region", "Amount"]}} =
               ExVEx.table(reopened, "Sales")
    end

    test "deleting every data row leaves a header and one blank row", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.put_table_totals_row(book, "Sales", Amount: :sum)
      {:ok, book} = ExVEx.delete_row(book, "Sheet1", 2, 3)
      reopened = round_trip(book, out)

      assert {:ok, %Table{ref: "A1:C2", totals_range: nil, columns: ["Region", "Qty", "Amount"]}} =
               ExVEx.table(reopened, "Sales")

      assert ExVEx.get_formula(reopened, "Sheet1", "C2") == {:ok, nil}
      assert ExVEx.get_cell(reopened, "Sheet1", "C2") == {:ok, nil}
    end

    test "deleting the header row adopts the next row's text as column names", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.delete_row(book, "Sheet1", 1)
      reopened = round_trip(book, out)

      assert {:ok, %Table{ref: "A1:C3", columns: ["North", "2", "10.5"]}} =
               ExVEx.table(reopened, "Sales")

      assert ExVEx.get_cell(reopened, "Sheet1", "B1") == {:ok, "2"}
    end

    test "deleting the whole table removes it from the package", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.delete_row(book, "Sheet1", 1, 4)
      reopened = round_trip(book, out)

      assert ExVEx.tables(reopened) == []
      refute Map.has_key?(reopened.parts, "xl/tables/table1.xml")
      refute part(reopened, "xl/worksheets/sheet1.xml") =~ "tableParts"
    end

    test "a shift on another sheet leaves the table alone", %{out: out} do
      {:ok, book} = ExVEx.add_table(sales_book(), "Sheet1", "A1:C4", name: "Sales")
      {:ok, book} = ExVEx.add_sheet(book, "Other")
      before = part(book, "xl/tables/table1.xml")
      {:ok, book} = ExVEx.insert_row(book, "Other", 1, 5)
      reopened = round_trip(book, out)
      assert part(reopened, "xl/tables/table1.xml") == before
    end
  end
end
