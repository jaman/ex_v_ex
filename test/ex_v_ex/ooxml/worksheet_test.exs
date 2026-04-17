defmodule ExVEx.OOXML.WorksheetTest do
  use ExUnit.Case, async: true

  alias ExVEx.OOXML.Worksheet, as: WorksheetXml
  alias ExVEx.OOXML.Worksheet.Cell
  alias ExVEx.Packaging.Zip
  alias ExVEx.Test.Fixtures

  @sheet_with_mixed_types """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
      <row r="1">
        <c r="A1" t="s"><v>0</v></c>
        <c r="B1"><v>42</v></c>
        <c r="C1" t="b"><v>1</v></c>
        <c r="D1" s="5" t="s"><v>3</v></c>
      </row>
      <row r="2">
        <c r="A2" t="inlineStr"><is><t>inline</t></is></c>
        <c r="B2"><v>3.14</v></c>
        <c r="C2" t="b"><v>0</v></c>
        <c r="E2" t="e"><v>#REF!</v></c>
      </row>
      <row r="3">
        <c r="A3" t="str"><f>=B1+B2</f><v>45.14</v></c>
        <c r="B3"><f>SUM(B1:B2)</f><v>45.14</v></c>
      </row>
    </sheetData>
  </worksheet>
  """

  describe "parse/1" do
    test "returns a map of coordinate => Cell covering every populated cell" do
      assert {:ok, %WorksheetXml{cells: cells}} = WorksheetXml.parse(@sheet_with_mixed_types)

      assert Map.has_key?(cells, {1, 1})
      assert Map.has_key?(cells, {2, 5})
      refute Map.has_key?(cells, {1, 5}), "empty cells should not be present"
    end

    test "parses shared string cells" do
      {:ok, %{cells: cells}} = WorksheetXml.parse(@sheet_with_mixed_types)

      assert %Cell{raw_type: :shared_string, raw_value: "0"} = Map.fetch!(cells, {1, 1})

      assert %Cell{raw_type: :shared_string, raw_value: "3", style_id: 5} =
               Map.fetch!(cells, {1, 4})
    end

    test "parses numeric cells (implicit type)" do
      {:ok, %{cells: cells}} = WorksheetXml.parse(@sheet_with_mixed_types)

      assert %Cell{raw_type: :number, raw_value: "42"} = Map.fetch!(cells, {1, 2})
      assert %Cell{raw_type: :number, raw_value: "3.14"} = Map.fetch!(cells, {2, 2})
    end

    test "parses boolean cells" do
      {:ok, %{cells: cells}} = WorksheetXml.parse(@sheet_with_mixed_types)

      assert %Cell{raw_type: :boolean, raw_value: "1"} = Map.fetch!(cells, {1, 3})
      assert %Cell{raw_type: :boolean, raw_value: "0"} = Map.fetch!(cells, {2, 3})
    end

    test "parses inline string cells" do
      {:ok, %{cells: cells}} = WorksheetXml.parse(@sheet_with_mixed_types)

      assert %Cell{raw_type: :inline_string, raw_value: "inline"} = Map.fetch!(cells, {2, 1})
    end

    test "parses error cells" do
      {:ok, %{cells: cells}} = WorksheetXml.parse(@sheet_with_mixed_types)

      assert %Cell{raw_type: :error, raw_value: "#REF!"} = Map.fetch!(cells, {2, 5})
    end

    test "captures formulas with cached values" do
      {:ok, %{cells: cells}} = WorksheetXml.parse(@sheet_with_mixed_types)

      assert %Cell{formula: "=B1+B2", raw_value: "45.14", raw_type: :formula_string} =
               Map.fetch!(cells, {3, 1})

      assert %Cell{formula: "SUM(B1:B2)", raw_value: "45.14", raw_type: :number} =
               Map.fetch!(cells, {3, 2})
    end

    test "parses the real cells.xlsx sheet1" do
      {:ok, entries} = Zip.read(Fixtures.path("cells.xlsx"))

      xml =
        entries
        |> Enum.find(&(&1.path == "xl/worksheets/sheet1.xml"))
        |> Map.fetch!(:data)

      assert {:ok, %WorksheetXml{cells: cells}} = WorksheetXml.parse(xml)

      assert %Cell{raw_type: :shared_string, raw_value: "0"} = Map.fetch!(cells, {1, 1})

      assert %Cell{raw_type: :shared_string, raw_value: "11", style_id: 3} =
               Map.fetch!(cells, {1, 3})
    end

    test "returns an error on malformed XML" do
      assert {:error, _} = WorksheetXml.parse("<worksheet></bad>")
    end
  end
end
