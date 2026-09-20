defmodule ExVEx.OOXML.Worksheet.EditableTablePartsTest do
  use ExUnit.Case, async: true

  alias ExVEx.OOXML.Worksheet
  alias ExVEx.OOXML.Worksheet.Editable

  @r_ns "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  defp editable(xml) do
    {:ok, tree} = Worksheet.parse_tree(xml)
    Editable.from_tree(tree)
  end

  defp encode(editable), do: editable |> Editable.to_tree() |> Worksheet.encode_tree()

  @bare ~s(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><pageMargins left="0.7"/><extLst><ext uri="x"/></extLst></worksheet>)

  @with_parts ~s(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="#{@r_ns}"><sheetData/><tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId3"/></tableParts></worksheet>)

  describe "table parts" do
    test "lists relationship ids" do
      assert Editable.table_part_rel_ids(editable(@with_parts)) == ["rId1", "rId3"]
      assert Editable.table_part_rel_ids(editable(@bare)) == []
    end

    test "adding to a sheet without tableParts inserts before extLst and declares the r namespace" do
      out =
        @bare
        |> editable()
        |> Editable.ensure_namespace("xmlns:r", @r_ns)
        |> Editable.add_table_part("rId7")
        |> encode()

      assert out =~ ~s(xmlns:r="#{@r_ns}")

      assert out =~
               ~s(<pageMargins left="0.7"/><tableParts count="1"><tablePart r:id="rId7"/></tableParts><extLst>)
    end

    test "adding to an existing tableParts appends and bumps the count" do
      out = @with_parts |> editable() |> Editable.add_table_part("rId9") |> encode()

      assert out =~
               ~s(<tableParts count="3"><tablePart r:id="rId1"/><tablePart r:id="rId3"/><tablePart r:id="rId9"/></tableParts>)
    end

    test "ensure_namespace is a no-op when already declared" do
      e = editable(@with_parts)
      assert Editable.ensure_namespace(e, "xmlns:r", @r_ns) === e
    end

    test "removing the last table part drops the element" do
      out =
        @with_parts
        |> editable()
        |> Editable.remove_table_part("rId1")
        |> Editable.remove_table_part("rId3")
        |> encode()

      refute out =~ "tableParts"
    end

    test "removing one keeps the others" do
      out = @with_parts |> editable() |> Editable.remove_table_part("rId1") |> encode()
      assert out =~ ~s(<tableParts count="1"><tablePart r:id="rId3"/></tableParts>)
    end
  end

  describe "move_cell/3" do
    test "moves the raw cell element and rewrites its coordinate" do
      e = editable(@bare)
      e = Editable.put_cell(e, {2, 1}, {:formula, "SUBTOTAL(109,T[a])"})
      e = Editable.move_cell(e, {2, 1}, {5, 1})

      assert Editable.get_cell(e, {2, 1}) == :error

      assert {:ok, {"c", [{"r", "A5"}], [{"f", [], ["SUBTOTAL(109,T[a])"]}]}} =
               Editable.get_cell(e, {5, 1})
    end

    test "moving an absent cell is a no-op" do
      e = editable(@bare)
      assert Editable.move_cell(e, {2, 1}, {5, 1}) === e
    end
  end
end
