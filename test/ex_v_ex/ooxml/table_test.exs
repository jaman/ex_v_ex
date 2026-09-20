defmodule ExVEx.OOXML.TableTest do
  use ExUnit.Case, async: true

  alias ExVEx.Mutation.Shift
  alias ExVEx.OOXML.Table
  alias ExVEx.OOXML.Table.Column
  alias ExVEx.Utils.Range

  @xml ~s(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>) <>
         ~s(<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3" id="1" name="T" displayName="T" ref="A1:C10" totalsRowShown="0">) <>
         ~s(<autoFilter ref="A1:C10"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter>) <>
         ~s(<sortState ref="A2:C10"><sortCondition ref="B2:B10"/></sortState>) <>
         ~s(<tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b" dataDxfId="3"/><tableColumn id="3" name="c"><calculatedColumnFormula>T[[#This Row],[a]]*2</calculatedColumnFormula></tableColumn></tableColumns>) <>
         ~s(<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>) <>
         ~s(<extLst><ext uri="{X}"><foo/></ext></extLst>) <>
         ~s(</table>)

  defp parse!(xml \\ @xml) do
    {:ok, table} = Table.parse(xml)
    table
  end

  describe "parse/1" do
    test "reads identity, ref, columns, and style" do
      table = parse!()
      assert table.id == 1
      assert table.name == "T"
      assert table.display_name == "T"
      assert table.ref == %Range{top_left: {1, 1}, bottom_right: {10, 3}}
      assert table.header_row_count == 1
      assert table.totals_row_count == 0
      assert table.totals_row_shown == false
      assert Enum.map(table.columns, & &1.name) == ["a", "b", "c"]
      assert Enum.map(table.columns, & &1.id) == [1, 2, 3]
      assert table.style.name == "TableStyleMedium2"
      assert table.style.show_row_stripes == true
      assert table.style.show_column_stripes == false
    end

    test "column formula is readable" do
      table = parse!()
      assert Column.calculated_formula(Enum.at(table.columns, 2)) == "T[[#This Row],[a]]*2"
      assert Column.calculated_formula(Enum.at(table.columns, 0)) == nil
    end

    test "rejects a non-table document" do
      assert {:error, :not_a_table_file} = Table.parse("<worksheet/>")
    end
  end

  describe "serialize/1" do
    test "an unmodified table serializes to the same bytes it was parsed from" do
      assert Table.serialize(parse!()) == @xml
    end

    test "unknown attributes and extension children survive a mutation" do
      out = parse!() |> Table.put_ref(Range.parse("A1:C12") |> elem(1)) |> Table.serialize()
      assert out =~ ~s(mc:Ignorable="xr xr3")
      assert out =~ ~s(dataDxfId="3")
      assert out =~ ~s(<extLst><ext uri="{X}"><foo/></ext></extLst>)
      assert out =~ ~s(ref="A1:C12")
      assert out =~ ~s(<autoFilter ref="A1:C12">)
    end

    test "new/1 builds a table Excel accepts" do
      table =
        Table.new(
          id: 4,
          name: "Sales",
          ref: Range.parse("B2:D5") |> elem(1),
          columns: ["Region", "Qty", "Amount"],
          style: %Table.StyleInfo{name: "TableStyleLight9"}
        )

      out = Table.serialize(table)

      assert out =~
               ~s(<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="4" name="Sales" displayName="Sales" ref="B2:D5" totalsRowShown="0">)

      assert out =~ ~s(<autoFilter ref="B2:D5"/>)

      assert out =~
               ~s(<tableColumns count="3"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Amount"/></tableColumns>)

      assert out =~
               ~s(<tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>)
    end
  end

  describe "shift/2 — rows" do
    test "insert above moves the whole table" do
      {:ok, table} = Table.shift(parse!(), Shift.insert(:row, 1, 2, nil))
      assert Range.to_string(table.ref) == "A3:C12"
      assert Table.auto_filter_ref(table) == "A3:C12"
      assert Table.serialize(table) =~ ~s(<sortState ref="A4:C12"><sortCondition ref="B4:B12"/>)
    end

    test "insert inside grows the table" do
      {:ok, table} = Table.shift(parse!(), Shift.insert(:row, 5, 2, nil))
      assert Range.to_string(table.ref) == "A1:C12"
    end

    test "insert below leaves it alone" do
      {:ok, table} = Table.shift(parse!(), Shift.insert(:row, 11, 2, nil))
      assert Range.to_string(table.ref) == "A1:C10"
    end

    test "delete covering the whole table removes it" do
      assert :deleted = Table.shift(parse!(), Shift.delete(:row, 1, 10, nil))
    end

    test "delete of some data rows shrinks the table" do
      {:ok, table} = Table.shift(parse!(), Shift.delete(:row, 3, 2, nil))
      assert Range.to_string(table.ref) == "A1:C8"
    end

    test "delete of every data row keeps a header and one blank data row" do
      {:ok, table} = Table.shift(parse!(), Shift.delete(:row, 2, 9, nil))
      assert Range.to_string(table.ref) == "A1:C2"
    end

    test "delete that covers the totals row drops it" do
      table = parse!() |> Table.put_totals_row_count(1)
      {:ok, shifted} = Table.shift(table, Shift.delete(:row, 10, 1, nil))
      assert shifted.totals_row_count == 0
      assert Range.to_string(shifted.ref) == "A1:C9"
    end

    test "delete of every data row with a totals row drops the totals row" do
      table = parse!() |> Table.put_totals_row_count(1)
      {:ok, shifted} = Table.shift(table, Shift.delete(:row, 2, 8, nil))
      assert shifted.totals_row_count == 0
      assert Range.to_string(shifted.ref) == "A1:C2"
    end
  end

  describe "shift/2 — columns" do
    test "insert inside adds columns with placeholder names at the insert position" do
      {:ok, table} = Table.shift(parse!(), Shift.insert(:col, 2, 2, nil))
      assert Range.to_string(table.ref) == "A1:E10"
      assert Enum.map(table.columns, & &1.name) == ["a", "Column2", "Column3", "b", "c"]
      assert Enum.map(table.columns, & &1.id) == [1, 4, 5, 2, 3]
    end

    test "insert inside shifts autoFilter column indexes" do
      {:ok, table} = Table.shift(parse!(), Shift.insert(:col, 2, 1, nil))
      assert Table.serialize(table) =~ ~s(<filterColumn colId="2">)
    end

    test "delete inside drops columns and their filters" do
      {:ok, table} = Table.shift(parse!(), Shift.delete(:col, 2, 1, nil))
      assert Range.to_string(table.ref) == "A1:B10"
      assert Enum.map(table.columns, & &1.name) == ["a", "c"]
      refute Table.serialize(table) =~ "filterColumn"
    end

    test "delete covering every column removes the table" do
      assert :deleted = Table.shift(parse!(), Shift.delete(:col, 1, 3, nil))
    end

    test "insert to the left moves the whole table" do
      {:ok, table} = Table.shift(parse!(), Shift.insert(:col, 1, 1, nil))
      assert Range.to_string(table.ref) == "B1:D10"
      assert Enum.map(table.columns, & &1.name) == ["a", "b", "c"]
    end
  end

  describe "column edits" do
    test "rename_column/3 updates the column name" do
      table = Table.rename_column(parse!(), "b", "Beta")
      assert Enum.map(table.columns, & &1.name) == ["a", "Beta", "c"]
    end

    test "put_columns/2 replaces the column list keeping ids unique" do
      table = Table.put_column_names(parse!(), ["x", "y", "z", "w"])
      assert Enum.map(table.columns, & &1.name) == ["x", "y", "z", "w"]
      assert Enum.map(table.columns, & &1.id) == [1, 2, 3, 4]
    end
  end
end
