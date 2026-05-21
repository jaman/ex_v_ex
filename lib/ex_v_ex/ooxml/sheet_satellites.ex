defmodule ExVEx.OOXML.SheetSatellites do
  @moduledoc """
  Cascades a `%ExVEx.Mutation.Shift{}` into the satellite parts linked
  from a worksheet's `.rels` file — comments, tables, and drawings.

  Each satellite part is a standalone XML document in the OOXML package.
  This module locates those parts via the worksheet's relationships, runs
  the appropriate shift module against their XML, and writes the result
  back into the workbook's `parts` map.
  """

  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.OOXML.{Comments, Drawing, Table}
  alias ExVEx.Packaging.Relationships
  alias ExVEx.Packaging.Relationships.Relationship
  alias ExVEx.Workbook

  @comments_rel "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  @table_rel "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
  @drawing_rel "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"

  @spec shift(Workbook.t(), String.t(), MutShift.t()) :: Workbook.t()
  def shift(%Workbook{} = book, sheet_path, %MutShift{} = mut_shift) do
    rels_path = rels_path_for(sheet_path)

    case Map.fetch(book.parts, rels_path) do
      {:ok, xml} ->
        case Relationships.parse(xml) do
          {:ok, rels} -> shift_linked_parts(book, rels, rels_path, mut_shift)
          _ -> book
        end

      :error ->
        book
    end
  end

  defp shift_linked_parts(%Workbook{} = book, %Relationships{entries: entries}, rels_path, shift) do
    Enum.reduce(entries, book, fn rel, acc ->
      shift_one(acc, rel, rels_path, shift)
    end)
  end

  defp shift_one(book, %Relationship{target_mode: :external}, _rels_path, _shift), do: book

  defp shift_one(book, %Relationship{type: @comments_rel} = rel, rels_path, shift) do
    apply_shift(book, Relationships.resolve(rel, rels_path), &Comments.shift_xml(&1, shift))
  end

  defp shift_one(book, %Relationship{type: @table_rel} = rel, rels_path, shift) do
    apply_shift(book, Relationships.resolve(rel, rels_path), &Table.shift_xml(&1, shift))
  end

  defp shift_one(book, %Relationship{type: @drawing_rel} = rel, rels_path, shift) do
    apply_shift(book, Relationships.resolve(rel, rels_path), &Drawing.shift_xml(&1, shift))
  end

  defp shift_one(book, _rel, _rels_path, _shift), do: book

  defp apply_shift(%Workbook{parts: parts} = book, part_path, shift_fun) do
    with {:ok, xml} <- Map.fetch(parts, part_path),
         {:ok, new_xml} <- shift_fun.(xml) do
      %{book | parts: Map.put(parts, part_path, new_xml)}
    else
      _ -> book
    end
  end

  defp rels_path_for(part_path) do
    dir = Path.dirname(part_path)
    base = Path.basename(part_path)

    case dir do
      "." -> "_rels/#{base}.rels"
      _ -> "#{dir}/_rels/#{base}.rels"
    end
  end
end
