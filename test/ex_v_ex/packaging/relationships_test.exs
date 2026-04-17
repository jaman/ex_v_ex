defmodule ExVEx.Packaging.RelationshipsTest do
  use ExUnit.Case, async: true

  alias ExVEx.Packaging.Relationships
  alias ExVEx.Packaging.Relationships.Relationship

  @workbook_rels """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>
  """

  describe "parse/1" do
    test "extracts every relationship with its id, type, and target" do
      assert {:ok, %Relationships{entries: entries}} = Relationships.parse(@workbook_rels)

      assert length(entries) == 3

      assert %Relationship{
               id: "rId1",
               type:
                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
               target: "worksheets/sheet1.xml"
             } in entries
    end

    test "returns an error for malformed XML" do
      assert {:error, _} = Relationships.parse("<Relationships></broken>")
    end
  end

  describe "get/2" do
    test "looks a relationship up by id" do
      {:ok, rels} = Relationships.parse(@workbook_rels)

      assert {:ok, %Relationship{target: "worksheets/sheet1.xml"}} =
               Relationships.get(rels, "rId1")
    end

    test "returns :error when the id is unknown" do
      {:ok, rels} = Relationships.parse(@workbook_rels)
      assert :error = Relationships.get(rels, "rId999")
    end
  end

  describe "resolve/2 — absolute package path of a relationship target" do
    test "joins the target against the source directory of the rels file" do
      {:ok, rels} = Relationships.parse(@workbook_rels)
      {:ok, rel} = Relationships.get(rels, "rId1")

      assert Relationships.resolve(rel, "xl/_rels/workbook.xml.rels") ==
               "xl/worksheets/sheet1.xml"
    end

    test "handles the package-root rels file (_rels/.rels)" do
      xml = """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>
      """

      {:ok, rels} = Relationships.parse(xml)
      {:ok, rel} = Relationships.get(rels, "rId1")

      assert Relationships.resolve(rel, "_rels/.rels") == "xl/workbook.xml"
    end
  end
end
