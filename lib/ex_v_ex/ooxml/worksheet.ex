defmodule ExVEx.OOXML.Worksheet do
  @moduledoc """
  Parser for `xl/worksheets/sheet*.xml`.

  Extracts the sparse cell grid from `<sheetData>` into a `%{coordinate =>
  Cell}` map. Surrounding worksheet elements (`<sheetViews>`, `<cols>`,
  `<mergeCells>`, `<pageMargins>`, etc.) are not yet modeled and are
  preserved at the byte level by the higher-level workbook.
  """

  alias ExVEx.OOXML.Worksheet.Cell
  alias ExVEx.Utils.Coordinate

  @type coordinate :: Coordinate.t()
  @type t :: %__MODULE__{cells: %{coordinate() => Cell.t()}}
  defstruct cells: %{}

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"worksheet", _attrs, children}} ->
        {:ok, %__MODULE__{cells: collect_cells(children)}}

      {:ok, _other} ->
        {:error, :not_a_worksheet}

      {:error, reason} ->
        {:error, reason}
    end
  end

  @doc """
  Surgically applies a cell mutation to a worksheet's raw XML and returns the
  new XML. Unrelated cells, rows, and surrounding worksheet content are
  preserved at the element level.

  Passing `nil` for value clears the cell.
  """
  @spec put_cell(binary(), coordinate(), ExVEx.cell_value()) ::
          {:ok, binary()} | {:error, term()}
  def put_cell(xml, coordinate, value) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"worksheet", attrs, children}} ->
        new_children = Enum.map(children, &mutate_sheet_data(&1, coordinate, value))
        tree = {"worksheet", attrs, new_children}
        {:ok, Saxy.encode!(tree, version: "1.0", encoding: "UTF-8", standalone: true)}

      {:ok, _other} ->
        {:error, :not_a_worksheet}

      {:error, reason} ->
        {:error, reason}
    end
  end

  defp mutate_sheet_data({"sheetData", attrs, rows}, coord, value) do
    {"sheetData", attrs, mutate_rows(rows, coord, value)}
  end

  defp mutate_sheet_data(other, _, _), do: other

  defp mutate_rows(rows, {row_num, _col} = coord, value) do
    case find_row_index(rows, row_num) do
      {:ok, idx} ->
        List.update_at(rows, idx, fn row -> mutate_row(row, coord, value) end)
        |> drop_empty_rows()

      :error ->
        case build_row(coord, value) do
          nil -> rows
          new_row -> insert_row(rows, new_row, row_num)
        end
    end
  end

  defp find_row_index(rows, row_num) do
    Enum.find_index(rows, fn
      {"row", attrs, _} -> row_attr(attrs) == row_num
      _ -> false
    end)
    |> case do
      nil -> :error
      idx -> {:ok, idx}
    end
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
