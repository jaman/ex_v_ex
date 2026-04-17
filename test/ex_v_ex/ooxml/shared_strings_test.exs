defmodule ExVEx.OOXML.SharedStringsTest do
  use ExUnit.Case, async: true

  alias ExVEx.OOXML.SharedStrings
  alias ExVEx.Packaging.Zip
  alias ExVEx.Test.Fixtures

  @plain_sst """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
    <si><t>Alpha</t></si>
    <si><t>Beta</t></si>
    <si><t>Gamma</t></si>
  </sst>
  """

  @rich_and_phonetic """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
    <si><t>plain</t><phoneticPr fontId="1"/></si>
    <si><r><rPr><b/></rPr><t>bold </t></r><r><t>tail</t></r></si>
  </sst>
  """

  describe "parse/1" do
    test "indexes plain strings by position" do
      assert {:ok, sst} = SharedStrings.parse(@plain_sst)

      assert SharedStrings.get(sst, 0) == {:ok, "Alpha"}
      assert SharedStrings.get(sst, 1) == {:ok, "Beta"}
      assert SharedStrings.get(sst, 2) == {:ok, "Gamma"}
      assert SharedStrings.count(sst) == 3
    end

    test "strips phonetic annotations and concatenates rich-text runs" do
      assert {:ok, sst} = SharedStrings.parse(@rich_and_phonetic)

      assert SharedStrings.get(sst, 0) == {:ok, "plain"}
      assert SharedStrings.get(sst, 1) == {:ok, "bold tail"}
    end

    test "handles the real cells.xlsx shared string table" do
      {:ok, entries} = Zip.read(Fixtures.path("cells.xlsx"))

      xml =
        entries
        |> Enum.find(&(&1.path == "xl/sharedStrings.xml"))
        |> Map.fetch!(:data)

      assert {:ok, sst} = SharedStrings.parse(xml)

      assert SharedStrings.get(sst, 0) == {:ok, "A1"}
      assert SharedStrings.get(sst, 1) == {:ok, "A2"}
      assert SharedStrings.get(sst, 2) == {:ok, "A3"}
      assert SharedStrings.get(sst, 3) == {:ok, "A4"}
    end

    test "get/2 returns :error for out-of-range indexes" do
      {:ok, sst} = SharedStrings.parse(@plain_sst)
      assert SharedStrings.get(sst, 99) == :error
      assert SharedStrings.get(sst, -1) == :error
    end

    test "returns an error on malformed XML" do
      assert {:error, _} = SharedStrings.parse("<sst><bad></sst>")
    end
  end
end
