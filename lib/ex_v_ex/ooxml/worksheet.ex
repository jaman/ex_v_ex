defmodule ExVEx.OOXML.Worksheet do
  @moduledoc """
  Parser and surgical mutator for `xl/worksheets/sheet*.xml`.

  Two sets of operations are exposed:

  - Binary-in / binary-out (`parse/1`, `put_cell/3`, `merge/3`, `unmerge/2`,
    `merged_ranges/1`) — convenient for one-off use but each pays the cost
    of a full `Saxy.SimpleForm` parse and re-emit.
  - Tree-in / tree-out (`parse_tree/1`, `encode_tree/1`,
    `put_cell_in_tree/3`, `merge_in_tree/3`, `unmerge_in_tree/2`,
    `merged_ranges_from_tree/1`, `cells_from_tree/1`) — used by
    `ExVEx.Workbook` to cache the parsed tree across many mutations and
    serialize only once at `ExVEx.save/2` time. This is what keeps bulk
    writes fast.

  Surrounding worksheet elements (`<sheetViews>`, `<cols>`,
  `<pageMargins>`, etc.) are not modeled; they pass through the
  SimpleForm round-trip unchanged.
  """

  alias ExVEx.OOXML.Worksheet.Cell
  alias ExVEx.Utils.{Coordinate, Range}

  @type coordinate :: Coordinate.t()
  @type tree :: {String.t(), list(), list()}
  @type t :: %__MODULE__{cells: %{coordinate() => Cell.t()}}
  defstruct cells: %{}

  @spec parse_tree(binary()) :: {:ok, tree()} | {:error, term()}
  def parse_tree(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"worksheet", _, _} = tree} -> {:ok, tree}
      {:ok, _other} -> {:error, :not_a_worksheet}
      {:error, reason} -> {:error, reason}
    end
  end

  @spec encode_tree(tree()) :: binary()
  def encode_tree({"worksheet", _, _} = tree) do
    Saxy.encode!(tree, version: "1.0", encoding: "UTF-8", standalone: true)
  end

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    with {:ok, tree} <- parse_tree(xml), do: {:ok, cells_from_tree(tree)}
  end

  @spec cells_from_tree(tree()) :: t()
  def cells_from_tree({"worksheet", _attrs, children}) do
    %__MODULE__{cells: collect_cells(children)}
  end

  @doc """
  Surgically applies a cell mutation to a worksheet's raw XML and returns
  the new XML. Unrelated cells, rows, and surrounding worksheet content
  are preserved at the element level.

  Passing `nil` for value clears the cell.
  """
  @spec put_cell(binary(), coordinate(), ExVEx.cell_value()) ::
          {:ok, binary()} | {:error, term()}
  def put_cell(xml, coordinate, value) when is_binary(xml) do
    with {:ok, tree} <- parse_tree(xml) do
      {:ok, encode_tree(put_cell_in_tree(tree, coordinate, value))}
    end
  end

  @spec put_cell_in_tree(tree(), coordinate(), ExVEx.cell_value()) :: tree()
  def put_cell_in_tree({"worksheet", attrs, children}, coordinate, value) do
    {"worksheet", attrs, Enum.map(children, &mutate_sheet_data(&1, coordinate, value))}
  end

  @doc """
  Reads the existing `<mergeCells>` list from a worksheet XML document.
  """
  @spec merged_ranges(binary()) :: {:ok, [Range.t()]} | {:error, term()}
  def merged_ranges(xml) when is_binary(xml) do
    with {:ok, tree} <- parse_tree(xml), do: {:ok, merged_ranges_from_tree(tree)}
  end

  @spec merged_ranges_from_tree(tree()) :: [Range.t()]
  def merged_ranges_from_tree({"worksheet", _attrs, children}),
    do: collect_merged_ranges(children)

  @doc """
  Adds a merged range to the worksheet. When `clear_non_anchor?` is true,
  every cell in the range other than the anchor (top-left) is removed from
  `<sheetData>` — Excel's visible convention.
  """
  @spec merge(binary(), Range.t(), boolean()) :: {:ok, binary()} | {:error, term()}
  def merge(xml, %Range{} = range, clear_non_anchor?) when is_binary(xml) do
    with {:ok, tree} <- parse_tree(xml) do
      {:ok, encode_tree(merge_in_tree(tree, range, clear_non_anchor?))}
    end
  end

  @spec merge_in_tree(tree(), Range.t(), boolean()) :: tree()
  def merge_in_tree({"worksheet", attrs, children}, %Range{} = range, clear_non_anchor?) do
    new_children =
      children
      |> maybe_clear_non_anchor(range, clear_non_anchor?)
      |> upsert_merge_cells_section(&add_merge_ref(&1, range))

    {"worksheet", attrs, new_children}
  end

  @doc "Removes a merged range from the worksheet."
  @spec unmerge(binary(), Range.t()) :: {:ok, binary()} | {:error, term()}
  def unmerge(xml, %Range{} = range) when is_binary(xml) do
    with {:ok, tree} <- parse_tree(xml),
         do: {:ok, encode_tree(unmerge_in_tree(tree, range))}
  end

  @spec unmerge_in_tree(tree(), Range.t()) :: tree()
  def unmerge_in_tree({"worksheet", attrs, children}, %Range{} = range) do
    new_children = upsert_merge_cells_section(children, &remove_merge_ref(&1, range))
    {"worksheet", attrs, new_children}
  end

  defp collect_merged_ranges(children) do
    children
    |> Enum.find_value([], fn
      {"mergeCells", _, items} -> items
      _ -> nil
    end)
    |> Enum.flat_map(&merge_cell_to_range/1)
  end

  defp merge_cell_to_range({"mergeCell", attrs, _}) do
    with {_, ref} <- List.keyfind(attrs, "ref", 0),
         {:ok, range} <- Range.parse(ref) do
      [range]
    else
      _ -> []
    end
  end

  defp merge_cell_to_range(_), do: []

  defp upsert_merge_cells_section(children, update_fn) do
    case Enum.find_index(children, &match?({"mergeCells", _, _}, &1)) do
      nil ->
        new_section = build_merge_cells([], update_fn)
        insert_merge_cells_section(children, new_section)

      idx ->
        {"mergeCells", _attrs, items} = Enum.at(children, idx)
        updated = build_merge_cells(items, update_fn)
        replace_or_drop(children, idx, updated)
    end
  end

  defp build_merge_cells(items, update_fn) do
    new_items = update_fn.(items)
    {"mergeCells", [{"count", Integer.to_string(length(new_items))}], new_items}
  end

  defp replace_or_drop(children, idx, {"mergeCells", _, []}), do: List.delete_at(children, idx)

  defp replace_or_drop(children, idx, new_section),
    do: List.replace_at(children, idx, new_section)

  defp insert_merge_cells_section(children, {"mergeCells", _, []}), do: children

  defp insert_merge_cells_section(children, new_section) do
    idx = Enum.find_index(children, &after_merge_cells?/1) || length(children)
    List.insert_at(children, idx, new_section)
  end

  defp after_merge_cells?({name, _, _})
       when name in [
              "phoneticPr",
              "conditionalFormatting",
              "dataValidations",
              "hyperlinks",
              "printOptions",
              "pageMargins",
              "pageSetup",
              "headerFooter",
              "rowBreaks",
              "colBreaks",
              "customProperties",
              "cellWatches",
              "ignoredErrors",
              "smartTags",
              "drawing",
              "legacyDrawing",
              "legacyDrawingHF",
              "picture",
              "oleObjects",
              "controls",
              "webPublishItems",
              "tableParts",
              "extLst"
            ],
       do: true

  defp after_merge_cells?(_), do: false

  defp add_merge_ref(items, %Range{} = range) do
    ref = Range.to_string(range)

    if Enum.any?(items, &merge_ref?(&1, ref)) do
      items
    else
      items ++ [{"mergeCell", [{"ref", ref}], []}]
    end
  end

  defp remove_merge_ref(items, %Range{} = range) do
    ref = Range.to_string(range)
    Enum.reject(items, &merge_ref?(&1, ref))
  end

  defp merge_ref?({"mergeCell", attrs, _}, ref) do
    case List.keyfind(attrs, "ref", 0) do
      {_, ^ref} -> true
      _ -> false
    end
  end

  defp merge_ref?(_, _), do: false

  defp maybe_clear_non_anchor(children, _range, false), do: children

  defp maybe_clear_non_anchor(children, range, true) do
    anchor = Range.anchor(range)

    Enum.map(children, fn
      {"sheetData", attrs, rows} -> {"sheetData", attrs, clear_rows(rows, range, anchor)}
      other -> other
    end)
  end

  defp clear_rows(rows, range, anchor) do
    rows
    |> Enum.map(fn
      {"row", attrs, cells} -> {"row", attrs, drop_non_anchor_cells(cells, range, anchor)}
      other -> other
    end)
    |> drop_empty_rows()
  end

  defp drop_non_anchor_cells(cells, range, anchor) do
    Enum.reject(cells, fn
      {"c", attrs, _} -> non_anchor_in_range?(attrs, range, anchor)
      _ -> false
    end)
  end

  defp non_anchor_in_range?(attrs, range, anchor) do
    with {_, ref} <- List.keyfind(attrs, "r", 0),
         {:ok, coord} <- Coordinate.parse(ref) do
      Range.contains?(range, coord) and coord != anchor
    else
      _ -> false
    end
  end

  defp mutate_sheet_data({"sheetData", attrs, rows}, coord, value) do
    {"sheetData", attrs, mutate_rows(rows, coord, value)}
  end

  defp mutate_sheet_data(other, _, _), do: other

  defp mutate_rows(rows, {row_num, _col} = coord, value) do
    case find_row_index(rows, row_num) do
      nil ->
        case build_row(coord, value) do
          nil -> rows
          new_row -> insert_row(rows, new_row, row_num)
        end

      idx ->
        rows
        |> List.update_at(idx, fn row -> mutate_row(row, coord, value) end)
        |> drop_empty_rows()
    end
  end

  defp find_row_index(rows, row_num) do
    Enum.find_index(rows, fn
      {"row", attrs, _} -> row_attr(attrs) == row_num
      _ -> false
    end)
  end

  defp mutate_row({"row", attrs, cells}, coord, value) do
    {"row", attrs, mutate_cells(cells, coord, value)}
  end

  defp mutate_cells(cells, coord, value) do
    ref = Coordinate.to_string(coord)

    case Enum.find_index(cells, &match_cell_ref?(&1, ref)) do
      nil ->
        case cell_element(coord, value, []) do
          nil -> cells
          new_cell -> insert_cell(cells, new_cell, coord)
        end

      idx ->
        existing = Enum.at(cells, idx)
        carry = carry_attrs(existing)

        case cell_element(coord, value, carry) do
          nil -> List.delete_at(cells, idx)
          new_cell -> List.replace_at(cells, idx, new_cell)
        end
    end
  end

  defp build_row({row_num, _} = coord, value) do
    case cell_element(coord, value, []) do
      nil -> nil
      cell -> {"row", [{"r", Integer.to_string(row_num)}], [cell]}
    end
  end

  defp insert_row(rows, new_row, row_num) do
    idx =
      Enum.find_index(rows, fn
        {"row", attrs, _} -> row_attr(attrs) > row_num
        _ -> false
      end) || length(rows)

    List.insert_at(rows, idx, new_row)
  end

  defp insert_cell(cells, new_cell, {_row, col}) do
    idx =
      Enum.find_index(cells, fn
        {"c", attrs, _} -> col_of_cell(attrs) > col
        _ -> false
      end) || length(cells)

    List.insert_at(cells, idx, new_cell)
  end

  defp match_cell_ref?({"c", attrs, _}, ref) do
    case List.keyfind(attrs, "r", 0) do
      {_, ^ref} -> true
      _ -> false
    end
  end

  defp match_cell_ref?(_, _), do: false

  defp row_attr(attrs) do
    case List.keyfind(attrs, "r", 0) do
      {_, value} -> String.to_integer(value)
      nil -> 0
    end
  end

  defp col_of_cell(attrs) do
    with {_, ref} <- List.keyfind(attrs, "r", 0),
         {:ok, {_, col}} <- Coordinate.parse(ref) do
      col
    else
      _ -> 0
    end
  end

  defp carry_attrs({"c", attrs, _}) do
    Enum.filter(attrs, fn {name, _} -> name == "s" end)
  end

  defp cell_element(_coord, nil, _carry), do: nil

  defp cell_element(coord, {:styled, inner, style_id}, carry)
       when is_integer(style_id) and style_id >= 0 do
    style_attr = {"s", Integer.to_string(style_id)}
    carry_without_s = Enum.reject(carry, fn {name, _} -> name == "s" end)
    cell_element(coord, inner, [style_attr | carry_without_s])
  end

  defp cell_element(coord, {:formula, formula}, carry) when is_binary(formula) do
    attrs = [{"r", Coordinate.to_string(coord)} | carry]
    {"c", attrs, [{"f", [], [formula]}]}
  end

  defp cell_element(coord, {:formula, formula, cached}, carry)
       when is_binary(formula) and (is_number(cached) or is_binary(cached) or is_boolean(cached)) do
    attrs = [{"r", Coordinate.to_string(coord)} | carry] ++ cached_type_attr(cached)

    {"c", attrs,
     [
       {"f", [], [formula]},
       {"v", [], [cached_to_string(cached)]}
     ]}
  end

  defp cell_element(coord, {:shared_string, index}, carry)
       when is_integer(index) and index >= 0 do
    attrs = [{"r", Coordinate.to_string(coord)} | carry] ++ [{"t", "s"}]
    {"c", attrs, [{"v", [], [Integer.to_string(index)]}]}
  end

  defp cell_element(coord, value, carry) when is_binary(value) do
    attrs = [{"r", Coordinate.to_string(coord)} | carry] ++ [{"t", "inlineStr"}]
    {"c", attrs, [{"is", [], [{"t", [], [value]}]}]}
  end

  defp cell_element(coord, value, carry) when is_boolean(value) do
    attrs = [{"r", Coordinate.to_string(coord)} | carry] ++ [{"t", "b"}]
    {"c", attrs, [{"v", [], [if(value, do: "1", else: "0")]}]}
  end

  defp cell_element(coord, value, carry) when is_integer(value) do
    attrs = [{"r", Coordinate.to_string(coord)} | carry]
    {"c", attrs, [{"v", [], [Integer.to_string(value)]}]}
  end

  defp cell_element(coord, value, carry) when is_float(value) do
    attrs = [{"r", Coordinate.to_string(coord)} | carry]
    {"c", attrs, [{"v", [], [float_to_string(value)]}]}
  end

  defp cached_type_attr(cached) when is_binary(cached), do: [{"t", "str"}]
  defp cached_type_attr(cached) when is_boolean(cached), do: [{"t", "b"}]
  defp cached_type_attr(cached) when is_number(cached), do: []

  defp cached_to_string(true), do: "1"
  defp cached_to_string(false), do: "0"
  defp cached_to_string(n) when is_integer(n), do: Integer.to_string(n)
  defp cached_to_string(n) when is_float(n), do: float_to_string(n)
  defp cached_to_string(bin) when is_binary(bin), do: bin

  defp float_to_string(value) do
    case Float.ratio(value) do
      {_, 1} -> :erlang.float_to_binary(value, [:compact, {:decimals, 1}])
      _ -> :erlang.float_to_binary(value, [:compact, {:decimals, 15}])
    end
  end

  defp drop_empty_rows(rows) do
    Enum.reject(rows, fn
      {"row", _attrs, []} -> true
      _ -> false
    end)
  end

  defp collect_cells(children) do
    children
    |> Enum.find_value([], fn
      {"sheetData", _, rows} -> rows
      _ -> nil
    end)
    |> Enum.reduce(%{}, &collect_row/2)
  end

  defp collect_row({"row", _attrs, cell_elements}, acc) do
    Enum.reduce(cell_elements, acc, &collect_cell/2)
  end

  defp collect_row(_, acc), do: acc

  defp collect_cell({"c", attrs, children}, acc) do
    case coordinate_from_attrs(attrs) do
      {:ok, coord} -> Map.put(acc, coord, build_cell(attrs, children))
      :error -> acc
    end
  end

  defp collect_cell(_, acc), do: acc

  defp coordinate_from_attrs(attrs) do
    case List.keyfind(attrs, "r", 0) do
      {_, ref} -> Coordinate.parse(ref)
      nil -> :error
    end
  end

  defp build_cell(attrs, children) do
    type_attr = type_attr(attrs)
    formula = find_child_text(children, "f")
    {raw_type, raw_value} = resolve_value(type_attr, formula, children)

    %Cell{
      raw_type: raw_type,
      raw_value: raw_value,
      formula: formula,
      style_id: style_id(attrs)
    }
  end

  defp type_attr(attrs) do
    case List.keyfind(attrs, "t", 0) do
      {_, value} -> value
      nil -> "n"
    end
  end

  defp style_id(attrs) do
    case List.keyfind(attrs, "s", 0) do
      {_, value} -> String.to_integer(value)
      nil -> nil
    end
  end

  defp resolve_value("inlineStr", _formula, children) do
    {:inline_string, inline_string_text(children)}
  end

  defp resolve_value("s", _formula, children),
    do: {:shared_string, find_child_text(children, "v")}

  defp resolve_value("b", _formula, children), do: {:boolean, find_child_text(children, "v")}
  defp resolve_value("e", _formula, children), do: {:error, find_child_text(children, "v")}

  defp resolve_value("str", formula, children) when is_binary(formula) do
    {:formula_string, find_child_text(children, "v")}
  end

  defp resolve_value(_number_like, _formula, children) do
    {:number, find_child_text(children, "v")}
  end

  defp inline_string_text(children) do
    case Enum.find(children, &match?({"is", _, _}, &1)) do
      {"is", _, is_children} -> text_from_rich_children(is_children)
      nil -> ""
    end
  end

  defp find_child_text(children, tag) do
    case Enum.find(children, &match?({^tag, _, _}, &1)) do
      {^tag, _, inner} -> text_content(inner)
      nil -> nil
    end
  end

  defp text_from_rich_children(children) do
    Enum.map_join(children, "", fn
      {"t", _, inner} -> text_content(inner)
      {"r", _, run_children} -> text_from_rich_children(run_children)
      _ -> ""
    end)
  end

  defp text_content(children) do
    Enum.map_join(children, "", fn
      text when is_binary(text) -> text
      _ -> ""
    end)
  end
end
