defmodule ExVEx.Packaging.ContentTypesTest do
  use ExUnit.Case, async: true

  alias ExVEx.Packaging.ContentTypes
  alias ExVEx.Packaging.ContentTypes.{Default, Override}
  alias ExVEx.Packaging.Zip
  alias ExVEx.Test.Fixtures

  describe "parse/1" do
    test "extracts defaults from a real [Content_Types].xml" do
      xml = fixture_content_types()
      assert {:ok, manifest} = ContentTypes.parse(xml)

      assert %Default{
               extension: "rels",
               content_type: "application/vnd.openxmlformats-package.relationships+xml"
             } in manifest.defaults

      assert %Default{extension: "xml", content_type: "application/xml"} in manifest.defaults
    end

    test "extracts overrides from a real [Content_Types].xml" do
      xml = fixture_content_types()
      assert {:ok, manifest} = ContentTypes.parse(xml)

      assert %Override{
               part_name: "/xl/workbook.xml",
               content_type:
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
             } in manifest.overrides
    end

    test "returns an error for malformed XML" do
      assert {:error, _} = ContentTypes.parse("<Types><broken></Types>")
    end
  end

  describe "serialize/1" do
    test "round-trips through parse/1 with equal contents" do
      {:ok, original} = ContentTypes.parse(fixture_content_types())
      serialized = ContentTypes.serialize(original)
      {:ok, re_parsed} = ContentTypes.parse(serialized)

      assert Enum.sort(original.defaults) == Enum.sort(re_parsed.defaults)
      assert Enum.sort(original.overrides) == Enum.sort(re_parsed.overrides)
    end

    test "produces well-formed XML with the OPC namespace" do
      manifest = %ContentTypes{
        defaults: [%Default{extension: "xml", content_type: "application/xml"}],
        overrides: []
      }

      xml = ContentTypes.serialize(manifest)
      assert xml =~ ~s(xmlns="http://schemas.openxmlformats.org/package/2006/content-types")
      assert xml =~ ~s(<?xml)
    end
  end

  defp fixture_content_types do
    {:ok, entries} = Zip.read(Fixtures.path("empty.xlsx"))

    Enum.find_value(entries, fn entry ->
      if entry.path == "[Content_Types].xml", do: entry.data
    end)
  end
end
