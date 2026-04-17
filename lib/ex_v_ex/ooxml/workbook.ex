defmodule ExVEx.OOXML.Workbook do
  @moduledoc """
  Parser for `xl/workbook.xml` — the per-workbook manifest of sheets.

  Only the parts ExVEx currently reasons about are modeled: the sheet list.
  Every other top-level element (`<workbookPr>`, `<bookViews>`, `<calcPr>`,
  `<definedNames>`, etc.) is preserved in the source bytes by the caller and
  round-trips untouched unless explicitly mutated later.
  """

  alias ExVEx.OOXML.Workbook.SheetRef

  @type t :: %__MODULE__{sheets: [SheetRef.t()]}
  defstruct sheets: []

  @rels_ns "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"workbook", _attrs, children}} ->
        {:ok, %__MODULE__{sheets: collect_sheets(children)}}

      {:ok, _other} ->
        {:error, :not_a_workbook}

      {:error, reason} ->
        {:error, reason}
    end
  end

  defp collect_sheets(children) do
    children
    |> Enum.find_value([], fn
      {"sheets", _, sheet_children} -> sheet_children
      _ -> nil
    end)
    |> Enum.flat_map(&sheet_ref_from_child/1)
  end

  defp sheet_ref_from_child({"sheet", attrs, _}) do
    [
      %SheetRef{
        name: fetch_attr!(attrs, "name"),
        sheet_id: attrs |> fetch_attr!("sheetId") |> String.to_integer(),
        rel_id: fetch_rel_id!(attrs),
        state: parse_state(attrs)
      }
    ]
  end

  defp sheet_ref_from_child(_), do: []

  defp fetch_rel_id!(attrs) do
    case List.keyfind(attrs, "r:id", 0) ||
           List.keyfind(attrs, "{#{@rels_ns}}id", 0) do
      {_, value} -> value
      nil -> raise ArgumentError, "missing r:id attribute on <sheet>"
    end
  end

  defp parse_state(attrs) do
    case List.keyfind(attrs, "state", 0) do
      {_, "hidden"} -> :hidden
      {_, "veryHidden"} -> :very_hidden
      _ -> :visible
    end
  end

  defp fetch_attr!(attrs, name) do
    case List.keyfind(attrs, name, 0) do
      {^name, value} -> value
      nil -> raise ArgumentError, "missing #{name} attribute on <sheet>"
    end
  end
end
