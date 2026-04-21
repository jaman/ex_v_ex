defmodule ExVEx.OOXML.Worksheet.Editable do
  @moduledoc """
  ETS-backed editable representation of a worksheet. The parsed SimpleForm
  tree is converted on first access and materialised back to a tree at
  serialise time.

  The cell grid lives in an ETS `:set` table for O(1) lookup and insert.
  Surrounding `<worksheet>` content (`<sheetViews>`, `<cols>`,
  `<mergeCells>`, `<pageMargins>`, …) passes through `pre_sheet_data` /
  `post_sheet_data` unchanged so fidelity is preserved.

  ## Sharing semantics

  ETS tables are mutable. Two `%ExVEx.Workbook{}` references that share
  the same cached sheet via `Workbook.put_sheet_tree/3` also share the
  same underlying ETS table; a mutation on one is visible from the other.

  In practice this shows up between a workbook and its `put_cell`
  descendants **after the sheet has been parsed at least once**. A pristine
  workbook from `ExVEx.open/1` that has not yet touched a sheet is
  insulated because it will lazily create a fresh table on first access.

  If you need an independent snapshot, `ExVEx.save/2` it and `open/1`
  that, or call `ExVEx.close/1` and re-open.

  ## Memory lifecycle

  ETS tables persist until the owning process exits. In short-lived
  processes (scripts, one-shot jobs, tests) this is fine. In long-running
  servers that open many workbooks, call `ExVEx.close/1` when done with a
  workbook to reclaim the tables eagerly.
  """

  alias ExVEx.Utils.{Coordinate, Range}

  @type cell_element :: {String.t(), list(), list()}
  @type t :: %__MODULE__{
          worksheet_attrs: list(),
          pre_sheet_data: list(),
          sheet_data_attrs: list(),
          row_attrs: %{pos_integer() => list()},
          cells_table: :ets.tid() | nil,
          post_sheet_data: list()
        }

  defstruct worksheet_attrs: [],
            pre_sheet_data: [],
            sheet_data_attrs: [],
            row_attrs: %{},
            cells_table: nil,
            post_sheet_data: []

  @spec from_tree({String.t(), list(), list()}) :: t()
  def from_tree({"worksheet", attrs, children}) do
    {pre, sheet_data, post} = split_on_sheet_data(children)

    {sd_attrs, row_attrs, cells} = extract_sheet_data(sheet_data)

    table = :ets.new(:exvex_cells, [:set, :public, {:write_concurrency, false}])
    Enum.each(cells, fn {coord, cell} -> :ets.insert(table, {coord, cell}) end)

    %__MODULE__{
      worksheet_attrs: attrs,
      pre_sheet_data: pre,
      sheet_data_attrs: sd_attrs,
      row_attrs: row_attrs,
      cells_table: table,
      post_sheet_data: post
    }
  end

  @spec to_tree(t()) :: {String.t(), list(), list()}
  def to_tree(%__MODULE__{} = e) do
    sheet_data = build_sheet_data(e)
    children = e.pre_sheet_data ++ [sheet_data] ++ e.post_sheet_data
    {"worksheet", e.worksheet_attrs, children}
  end

  @spec put_cell(t(), ExVEx.Utils.Coordinate.t(), ExVEx.cell_value()) :: t()
  def put_cell(%__MODULE__{cells_table: table, row_attrs: row_attrs} = e, coord, value) do
    existing_carry =
      case :ets.lookup(table, coord) do
        [{^coord, {"c", attrs, _}}] -> Enum.filter(attrs, fn {name, _} -> name == "s" end)
        [] -> []
      end

    case cell_element(coord, value, existing_carry) do
      nil ->
        :ets.delete(table, coord)
        e

      new_cell ->
        :ets.insert(table, {coord, new_cell})
        new_row_attrs = ensure_row_attrs(row_attrs, coord)
        if new_row_attrs === row_attrs, do: e, else: %{e | row_attrs: new_row_attrs}
    end
  end

  @spec get_cell(t(), ExVEx.Utils.Coordinate.t()) :: {:ok, cell_element()} | :error
  def get_cell(%__MODULE__{cells_table: table}, coord) do
    case :ets.lookup(table, coord) do
      [{^coord, cell}] -> {:ok, cell}
      [] -> :error
    end
  end

  @doc """
  Returns the resolved `%ExVEx.OOXML.Worksheet.Cell{}` record for the given
  coordinate, or `:error` if the cell is absent.
  """
  @spec cell_record_at(t(), ExVEx.Utils.Coordinate.t()) ::
          {:ok, ExVEx.OOXML.Worksheet.Cell.t()} | :error
  def cell_record_at(%__MODULE__{cells_table: table}, coord) do
    case :ets.lookup(table, coord) do
      [{^coord, cell_element}] -> {:ok, build_cell_record(cell_element)}
      [] -> :error
    end
  end

  @spec cells_map(t()) :: %{ExVEx.Utils.Coordinate.t() => ExVEx.OOXML.Worksheet.Cell.t()}
  def cells_map(%__MODULE__{cells_table: table}) do
    table
    |> :ets.tab2list()
    |> Map.new(fn {coord, cell_element} -> {coord, build_cell_record(cell_element)} end)
  end

  @spec merged_ranges(t()) :: [Range.t()]
  def merged_ranges(%__MODULE__{post_sheet_data: post}) do
    case find_merge_cells_section(post) do
      nil -> []
      {"mergeCells", _, items} -> Enum.flat_map(items, &merge_cell_to_range/1)
    end
  end

  @spec merge(t(), Range.t(), boolean()) :: t()
  def merge(%__MODULE__{cells_table: table} = e, %Range{} = range, clear_non_anchor?) do
    if clear_non_anchor?, do: clear_non_anchor_cells(table, range)

    new_post = upsert_merge_cells(e.post_sheet_data, fn items -> add_merge_ref(items, range) end)
    %{e | post_sheet_data: new_post}
  end

  defp clear_non_anchor_cells(table, range) do
    anchor = Range.anchor(range)
    Enum.each(:ets.tab2list(table), &maybe_delete_cell(&1, table, range, anchor))
  end

  defp maybe_delete_cell({coord, _}, table, range, anchor) do
    if coord != anchor and Range.contains?(range, coord), do: :ets.delete(table, coord)
  end

  @spec unmerge(t(), Range.t()) :: t()
  def unmerge(%__MODULE__{} = e, %Range{} = range) do
    new_post =
      upsert_merge_cells(e.post_sheet_data, fn items -> remove_merge_ref(items, range) end)

    %{e | post_sheet_data: new_post}
  end

  defp split_on_sheet_data(children) do
    idx =
      Enum.find_index(children, fn
        {"sheetData", _, _} -> true
        _ -> false
      end)

    case idx do
      nil ->
        {children, nil, []}

      _ ->
        {Enum.slice(children, 0, idx), Enum.at(children, idx),
         Enum.slice(children, (idx + 1)..-1//1)}
    end
  end

  defp extract_sheet_data(nil), do: {[], %{}, %{}}

  defp extract_sheet_data({"sheetData", attrs, rows}) do
    {row_attrs, cells} =
      Enum.reduce(rows, {%{}, %{}}, fn
        {"row", row_attrs_list, cell_elems}, {ra, cs} ->
          row_num = row_num_from_attrs(row_attrs_list)
          new_ra = if row_num, do: Map.put(ra, row_num, row_attrs_list), else: ra

          new_cs = Enum.reduce(cell_elems, cs, &collect_cell_into_map/2)

          {new_ra, new_cs}

        _, acc ->
          acc
      end)

    {attrs, row_attrs, cells}
  end

  defp collect_cell_into_map({"c", c_attrs, _} = cell, acc) do
    case coord_from_cell_attrs(c_attrs) do
      {:ok, coord} -> Map.put(acc, coord, cell)
      :error -> acc
    end
  end

  defp collect_cell_into_map(_, acc), do: acc

  defp row_num_from_attrs(attrs) do
    case List.keyfind(attrs, "r", 0) do
      {_, value} ->
        case Integer.parse(value) do
          {n, _} -> n
          _ -> nil
        end

      nil ->
        nil
    end
  end

  defp coord_from_cell_attrs(attrs) do
    case List.keyfind(attrs, "r", 0) do
      {_, ref} -> Coordinate.parse(ref)
      nil -> :error
    end
  end

  defp ensure_row_attrs(row_attrs, {row, _col}) do
    case Map.has_key?(row_attrs, row) do
      true -> row_attrs
      false -> Map.put(row_attrs, row, [{"r", Integer.to_string(row)}])
    end
  end

  defp build_sheet_data(%__MODULE__{cells_table: table} = e) do
    cells_by_row = group_cells_by_row(:ets.tab2list(table))
    rows_to_emit = rows_to_emit(e.row_attrs, cells_by_row)
    row_tuples = Enum.map(rows_to_emit, &build_row(&1, e.row_attrs, cells_by_row))
    {"sheetData", e.sheet_data_attrs, row_tuples}
  end

  defp group_cells_by_row(cells_list) do
    Enum.reduce(cells_list, %{}, fn {{row, _} = coord, cell}, acc ->
      Map.update(acc, row, [{coord, cell}], &[{coord, cell} | &1])
    end)
  end

  defp rows_to_emit(row_attrs, cells_by_row) do
    attr_rows = Map.keys(row_attrs)
    cell_rows = Map.keys(cells_by_row)
    (attr_rows ++ cell_rows) |> Enum.uniq() |> Enum.sort()
  end

  defp build_row(row_num, row_attrs_map, cells_by_row) do
    attrs = Map.get(row_attrs_map, row_num, [{"r", Integer.to_string(row_num)}])

    cells =
      case Map.get(cells_by_row, row_num, []) do
        [] ->
          []

        entries ->
          entries
          |> Enum.sort_by(fn {{_, col}, _} -> col end)
          |> Enum.map(fn {_, cell} -> cell end)
      end

    {"row", attrs, cells}
  end

  defp find_merge_cells_section(children) do
    Enum.find(children, &match?({"mergeCells", _, _}, &1))
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

  defp upsert_merge_cells(children, update_fn) do
    case Enum.find_index(children, &match?({"mergeCells", _, _}, &1)) do
      nil ->
        new_section = build_merge_cells_element([], update_fn)
        insert_merge_cells_element(children, new_section)

      idx ->
        {"mergeCells", _attrs, items} = Enum.at(children, idx)
        new_section = build_merge_cells_element(items, update_fn)
        replace_or_drop(children, idx, new_section)
    end
  end

  defp build_merge_cells_element(items, update_fn) do
    new_items = update_fn.(items)
    {"mergeCells", [{"count", Integer.to_string(length(new_items))}], new_items}
  end

  defp replace_or_drop(children, idx, {"mergeCells", _, []}), do: List.delete_at(children, idx)

  defp replace_or_drop(children, idx, new_section),
    do: List.replace_at(children, idx, new_section)

  defp insert_merge_cells_element(children, {"mergeCells", _, []}), do: children

  defp insert_merge_cells_element(children, new_section) do
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

  defp build_cell_record({"c", attrs, children}) do
    alias ExVEx.OOXML.Worksheet.Cell

    type_attr =
      case List.keyfind(attrs, "t", 0) do
        {_, v} -> v
        nil -> "n"
      end

    style_id =
      case List.keyfind(attrs, "s", 0) do
        {_, v} -> String.to_integer(v)
        nil -> nil
      end

    formula = find_child_text(children, "f")
    {raw_type, raw_value} = resolve_value(type_attr, formula, children)

    %Cell{
      raw_type: raw_type,
      raw_value: raw_value,
      formula: formula,
      style_id: style_id
    }
  end

  defp resolve_value("inlineStr", _formula, children),
    do: {:inline_string, inline_string_text(children)}

  defp resolve_value("s", _formula, children),
    do: {:shared_string, find_child_text(children, "v")}

  defp resolve_value("b", _formula, children), do: {:boolean, find_child_text(children, "v")}
  defp resolve_value("e", _formula, children), do: {:error, find_child_text(children, "v")}

  defp resolve_value("str", formula, children) when is_binary(formula),
    do: {:formula_string, find_child_text(children, "v")}

  defp resolve_value(_number_like, _formula, children),
    do: {:number, find_child_text(children, "v")}

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
end
