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

  @doc """
  Rewrites `xl/workbook.xml` so that its `<calcPr>` element carries
  `fullCalcOnLoad="1"` — forcing Excel to recompute every formula on open.
  Creates a `<calcPr>` element if one isn't present. All other worksheet
  content is preserved at the element level via Saxy SimpleForm round-trip.
  """
  @spec set_full_calc_on_load(binary()) :: {:ok, binary()} | {:error, term()}
  def set_full_calc_on_load(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"workbook", attrs, children}} ->
        tree = {"workbook", attrs, ensure_full_calc(children)}
        {:ok, Saxy.encode!(tree, version: "1.0", encoding: "UTF-8", standalone: true)}

      {:ok, _other} ->
        {:error, :not_a_workbook}

      {:error, reason} ->
        {:error, reason}
    end
  end

  defp ensure_full_calc(children) do
    case Enum.find_index(children, &match?({"calcPr", _, _}, &1)) do
      nil ->
        children ++ [{"calcPr", [{"fullCalcOnLoad", "1"}], []}]

      idx ->
        {"calcPr", pr_attrs, pr_children} = Enum.at(children, idx)
        new_attrs = upsert_attr(pr_attrs, "fullCalcOnLoad", "1")
        List.replace_at(children, idx, {"calcPr", new_attrs, pr_children})
    end
  end

  defp upsert_attr(attrs, name, value) do
    case List.keyfind(attrs, name, 0) do
      nil -> attrs ++ [{name, value}]
      _ -> List.keyreplace(attrs, name, 0, {name, value})
    end
  end

  @doc """
  Rewrites the `<sheets>` section of `xl/workbook.xml` using the sheet list
  from `workbook`, preserving all other elements at the element level.
  """
  @spec serialize_into(t(), binary()) :: binary()
  def serialize_into(%__MODULE__{sheets: sheets}, original_xml) when is_binary(original_xml) do
    case Saxy.SimpleForm.parse_string(original_xml) do
      {:ok, {"workbook", attrs, children}} ->
        new_children = replace_sheets(children, sheets)
        tree = {"workbook", attrs, new_children}
        Saxy.encode!(tree, version: "1.0", encoding: "UTF-8", standalone: true)

      _ ->
        original_xml
    end
  end

  defp replace_sheets(children, sheets) do
    new_section = sheets_element(sheets)

    case Enum.find_index(children, &match?({"sheets", _, _}, &1)) do
      nil -> insert_sheets_element(children, new_section)
      idx -> List.replace_at(children, idx, new_section)
    end
  end

  defp sheets_element(sheets) do
    {"sheets", [], Enum.map(sheets, &sheet_element/1)}
  end

  defp sheet_element(%SheetRef{name: name, sheet_id: id, rel_id: rel_id, state: state}) do
    attrs =
      [
        {"name", name},
        {"sheetId", Integer.to_string(id)},
        {"r:id", rel_id}
      ] ++ state_attrs(state)

    {"sheet", attrs, []}
  end

  defp state_attrs(:hidden), do: [{"state", "hidden"}]
  defp state_attrs(:very_hidden), do: [{"state", "veryHidden"}]
  defp state_attrs(_), do: []

  defp insert_sheets_element(children, new_section) do
    idx =
      Enum.find_index(children, fn
        {"calcPr", _, _} -> true
        {"definedNames", _, _} -> true
        {"workbookProtection", _, _} -> true
        _ -> false
      end) || length(children)

    List.insert_at(children, idx, new_section)
  end

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
