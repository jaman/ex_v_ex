defmodule ExVEx.TableFixtureTest do
  use ExUnit.Case, async: true

  alias ExVEx.Table
  alias ExVEx.Test.Fixtures

  @fixture Fixtures.path("with_table.xlsx")
  @table_part "xl/tables/table1.xml"
  @sales_part "xl/worksheets/sheet1.xml"
  @summary_part "xl/worksheets/sheet2.xml"

  setup do
    out = Fixtures.tmp_path("with_table_out.xlsx")
    on_exit(fn -> File.rm(out) end)
    {:ok, book} = ExVEx.open(@fixture)
    %{out: out, book: book, original: book.parts}
  end

  defp round_trip(book, out) do
    :ok = ExVEx.save(book, out)
    {:ok, reopened} = ExVEx.open(out)
    reopened
  end

  test "an untouched workbook round-trips every part byte for byte", %{
    book: book,
    out: out,
    original: original
  } do
    reopened = round_trip(book, out)
    assert reopened.parts == original
  end

  test "reading the table changes nothing", %{book: book, out: out, original: original} do
    assert {:ok, %Table{name: "Orders", ref: "A1:D5", totals_range: "A5:D5"}} =
             ExVEx.table(book, "Orders")

    assert {:ok, [["North", 2, 10.5, 21] | _]} = ExVEx.table_rows(book, "Orders")
    assert round_trip(book, out).parts == original
  end

  test "a write on another sheet leaves the table and its sheet untouched", %{
    book: book,
    out: out,
    original: original
  } do
    {:ok, book} = ExVEx.put_cell(book, "Summary", "A3", "note")
    reopened = round_trip(book, out)

    assert reopened.parts[@table_part] == original[@table_part]
    assert reopened.parts[@sales_part] == original[@sales_part]
  end

  test "a shift on another sheet leaves the table and its sheet untouched", %{
    book: book,
    out: out,
    original: original
  } do
    {:ok, book} = ExVEx.insert_row(book, "Summary", 1, 3)
    reopened = round_trip(book, out)

    assert reopened.parts[@table_part] == original[@table_part]
    assert reopened.parts[@sales_part] == original[@sales_part]
    assert ExVEx.get_formula(reopened, "Summary", "B4") == {:ok, "SUM(Orders[Amount])"}
  end

  test "inserting rows inside the table grows it and leaves every formula alone", %{
    book: book,
    out: out,
    original: original
  } do
    {:ok, book} = ExVEx.insert_row(book, "Sales", 3, 2)
    reopened = round_trip(book, out)

    assert {:ok, %Table{ref: "A1:D7", data_range: "A2:D6", totals_range: "A7:D7"}} =
             ExVEx.table(reopened, "Orders")

    assert ExVEx.get_formula(reopened, "Sales", "D5") ==
             {:ok, "Orders[[#This Row],[Qty]]*Orders[[#This Row],[Amount]]"}

    assert ExVEx.get_formula(reopened, "Sales", "C7") == {:ok, "SUBTOTAL(109,Orders[Amount])"}
    assert reopened.parts[@summary_part] == original[@summary_part]
    assert reopened.parts[@table_part] =~ ~s(xr:uid="{00000000-000C-0000-FFFF-FFFF00000000}")
    assert reopened.parts[@table_part] =~ ~s(<autoFilter ref="A1:D6" xr:uid=)

    assert reopened.parts[@table_part] =~
             ~s(<calculatedColumnFormula>Orders[[#This Row],[Qty]]*Orders[[#This Row],[Amount]]</calculatedColumnFormula>)
  end

  test "renaming a column rewrites every formula that names it", %{book: book, out: out} do
    {:ok, book} = ExVEx.rename_table_column(book, "Orders", "Amount", "Value")
    reopened = round_trip(book, out)

    assert ExVEx.get_cell(reopened, "Sales", "C1") == {:ok, "Value"}

    assert ExVEx.get_formula(reopened, "Sales", "D2") ==
             {:ok, "Orders[[#This Row],[Qty]]*Orders[[#This Row],[Value]]"}

    assert ExVEx.get_formula(reopened, "Sales", "C5") == {:ok, "SUBTOTAL(109,Orders[Value])"}
    assert ExVEx.get_formula(reopened, "Summary", "B1") == {:ok, "SUM(Orders[Value])"}
    assert ExVEx.get_formula(reopened, "Summary", "B2") == {:ok, "COUNTA(Orders[Region])"}

    assert [%{name: "GrandTotal", reference: "SUM(Orders[Value])"}] =
             ExVEx.defined_names(reopened)

    assert reopened.parts[@table_part] =~
             ~s(<tableColumn id="3" xr3:uid="{00000000-0010-0000-0100-000003000000}" name="Value" totalsRowFunction="sum"/>)

    assert reopened.parts[@table_part] =~
             ~s(<calculatedColumnFormula>Orders[[#This Row],[Qty]]*Orders[[#This Row],[Value]]</calculatedColumnFormula>)
  end

  test "appending rows fills calculated columns and moves the totals row", %{book: book, out: out} do
    {:ok, book} = ExVEx.append_table_rows(book, "Orders", [["West", 4, 2.0, nil]])
    reopened = round_trip(book, out)

    assert {:ok, %Table{ref: "A1:D6", data_range: "A2:D5", totals_range: "A6:D6"}} =
             ExVEx.table(reopened, "Orders")

    assert ExVEx.get_cell(reopened, "Sales", "A5") == {:ok, "West"}

    assert ExVEx.get_formula(reopened, "Sales", "D5") ==
             {:ok, "Orders[[#This Row],[Qty]]*Orders[[#This Row],[Amount]]"}

    assert ExVEx.get_cell(reopened, "Sales", "A6") == {:ok, "Total"}
    assert ExVEx.get_formula(reopened, "Sales", "C6") == {:ok, "SUBTOTAL(109,Orders[Amount])"}
  end

  test "removing the table converts every structured reference", %{book: book, out: out} do
    {:ok, book} = ExVEx.remove_table(book, "Orders")
    reopened = round_trip(book, out)

    assert ExVEx.tables(reopened) == []
    assert ExVEx.get_formula(reopened, "Sales", "D3") == {:ok, "B3*C3"}
    assert ExVEx.get_formula(reopened, "Sales", "C5") == {:ok, "SUBTOTAL(109,C2:C4)"}
    assert ExVEx.get_formula(reopened, "Summary", "B1") == {:ok, "SUM(Sales!C2:C4)"}
    assert [%{name: "GrandTotal", reference: "SUM(Sales!C2:C4)"}] = ExVEx.defined_names(reopened)
    refute Map.has_key?(reopened.parts, @table_part)
  end
end
