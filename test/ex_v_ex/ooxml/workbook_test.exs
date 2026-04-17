defmodule ExVEx.OOXML.WorkbookTest do
  use ExUnit.Case, async: true

  alias ExVEx.OOXML.Workbook, as: WorkbookXml
  alias ExVEx.OOXML.Workbook.SheetRef

  @workbook_xml """
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
      <sheet name="Inventory" sheetId="1" r:id="rId1"/>
      <sheet name="Pricing" sheetId="2" r:id="rId2"/>
      <sheet name="Archive" sheetId="3" r:id="rId3" state="hidden"/>
    </sheets>
  </workbook>
  """

  describe "parse/1" do
    test "extracts every sheet with its name, internal id, and relationship id" do
      assert {:ok, %WorkbookXml{sheets: sheets}} = WorkbookXml.parse(@workbook_xml)

      assert [
               %SheetRef{name: "Inventory", sheet_id: 1, rel_id: "rId1", state: :visible},
               %SheetRef{name: "Pricing", sheet_id: 2, rel_id: "rId2", state: :visible},
               %SheetRef{name: "Archive", sheet_id: 3, rel_id: "rId3", state: :hidden}
             ] = sheets
    end

    test "returns an error for malformed XML" do
      assert {:error, _} = WorkbookXml.parse("<workbook></broken>")
    end
  end
end
