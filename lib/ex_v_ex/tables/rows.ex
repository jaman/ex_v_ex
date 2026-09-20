defmodule ExVEx.Tables.Rows do
  @moduledoc """
  Row-level table operations: reading the data body, appending rows,
  and adding, updating, or removing the totals row.

  Appending rows extends the table downward. The cells that would be
  written must be empty; a totals row is moved down to stay below the
  data. Columns with a calculated-column formula receive that formula in
  new rows unless the row supplies a value for them.
  """

  alias ExVEx.OOXML.Table
  alias ExVEx.OOXML.Table.Column
  alias ExVEx.OOXML.Worksheet.Editable
  alias ExVEx.Tables
  alias ExVEx.Tables.{Cells, Validation}
  alias ExVEx.Utils.Range
  alias ExVEx.Workbook
  alias ExVEx.Workbook.TableParts
  alias ExVEx.Workbook.TableParts.Entry

  @type totals_spec ::
          Column.totals_function() | {:custom, String.t()} | {:label, String.t()}

  @spec rows(Workbook.t(), String.t()) :: {:ok, [[term()]]} | {:error, term()}
  def rows(%Workbook{} = book, name) do
    with {:ok, entry} <- Tables.fetch_entry(book, name) do
      {:ok, read_rows(book, entry)}
    end
  end

  @spec records(Workbook.t(), String.t()) :: {:ok, [%{String.t() => term()}]} | {:error, term()}
  def records(%Workbook{} = book, name) do
    with {:ok, entry} <- Tables.fetch_entry(book, name) do
      names = Table.column_names(entry.table)
      {:ok, book |> read_rows(entry) |> Enum.map(&Map.new(Enum.zip(names, &1)))}
    end
  end

  @spec append(Workbook.t(), String.t(), [[term()] | %{String.t() => term()}]) :: Tables.result()
  def append(%Workbook{} = book, name, rows) when is_list(rows) do
    with {:ok, entry} <- Tables.fetch_entry(book, name),
         {:ok, lists} <- normalize_rows(entry.table, rows) do
      append_lists(book, entry, lists)
    end
  end

  @spec put_totals(
          Workbook.t(),
          String.t(),
          keyword() | %{optional(String.t() | atom()) => totals_spec()}
        ) ::
          Tables.result()
  def put_totals(%Workbook{} = book, name, functions) do
    specs = Enum.map(functions, fn {column, spec} -> {to_string(column), spec} end)

    with {:ok, entry} <- Tables.fetch_entry(book, name),
         :ok <- check_columns(entry.table, Enum.map(specs, &elem(&1, 0))),
         :ok <- check_specs(specs),
         {:ok, table, book} <- ensure_totals_row(book, entry) do
      table =
        Enum.reduce(specs, table, fn {column, spec}, acc ->
          put_column_totals(acc, column, spec)
        end)

      book = TableParts.put(book, entry, table)
      {:ok, Enum.reduce(specs, book, &write_totals_cell(&2, entry.sheet_path, table, &1))}
    end
  end

  @spec remove_totals(Workbook.t(), String.t()) :: Tables.result()
  def remove_totals(%Workbook{} = book, name) do
    with {:ok, entry} <- Tables.fetch_entry(book, name) do
      remove_totals_row(book, entry)
    end
  end

  defp read_rows(book, %Entry{table: table} = entry) do
    {_, left} = table.ref.top_left
    {_, right} = table.ref.bottom_right

    case Table.data_rows(table) do
      nil ->
        []

      {first, last} ->
        Enum.map(first..last//1, &read_row(book, entry.sheet_path, &1, left, right))
    end
  end

  defp read_row(book, sheet_path, row, left, right) do
    Enum.map(left..right//1, fn col ->
      {:ok, value, _book} = Cells.read(book, sheet_path, {row, col})
      value
    end)
  end

  defp normalize_rows(%Table{} = table, rows) do
    names = Table.column_names(table)
    width = length(names)

    Enum.reduce_while(rows, {:ok, []}, fn row, {:ok, acc} ->
      case normalize_row(row, names, width) do
        {:ok, list} -> {:cont, {:ok, [list | acc]}}
        {:error, _} = error -> {:halt, error}
      end
    end)
    |> case do
      {:ok, lists} -> {:ok, Enum.reverse(lists)}
      error -> error
    end
  end

  defp normalize_row(row, _names, width) when is_list(row) do
    if length(row) == width, do: {:ok, row}, else: {:error, :column_count_mismatch}
  end

  defp normalize_row(row, names, _width) when is_map(row) do
    lookup = Map.new(row, fn {key, value} -> {String.downcase(to_string(key)), value} end)
    known = MapSet.new(names, &String.downcase/1)

    case Enum.find(Map.keys(lookup), &(not MapSet.member?(known, &1))) do
      nil -> {:ok, Enum.map(names, &Map.get(lookup, String.downcase(&1)))}
      unknown -> {:error, {:unknown_column, original_key(row, unknown)}}
    end
  end

  defp normalize_row(_row, _names, _width), do: {:error, :invalid_row}

  defp original_key(row, downcased) do
    row |> Map.keys() |> Enum.map(&to_string/1) |> Enum.find(&(String.downcase(&1) == downcased))
  end

  defp append_lists(book, _entry, []), do: {:ok, book}

  defp append_lists(book, %Entry{table: table} = entry, lists) do
    count = length(lists)
    {_, left} = table.ref.top_left
    {bottom, right} = table.ref.bottom_right
    first_new = bottom - table.totals_row_count + 1
    target = %Range{top_left: {first_new, left}, bottom_right: {first_new + count - 1, right}}
    new_bottom = bottom + count

    with {:ok, book} <- check_target_empty(book, entry, target, new_bottom) do
      book =
        book
        |> move_totals_cells(entry, bottom, new_bottom)
        |> write_rows(entry, lists, first_new)

      new_ref = %{table.ref | bottom_right: {new_bottom, right}}
      {:ok, TableParts.put(book, entry, Table.put_ref(table, new_ref))}
    end
  end

  defp check_target_empty(book, %Entry{table: table} = entry, target, new_bottom) do
    {_, left} = table.ref.top_left
    {bottom, right} = table.ref.bottom_right
    totals_row = Table.totals_row(table)

    cells_to_check =
      Range.cells(target) ++ totals_destination(totals_row, new_bottom, left, right)

    {:ok, editable, book} = Workbook.fetch_sheet_tree(book, entry.sheet_path)

    occupied =
      Enum.find(cells_to_check, fn {row, _} = coord ->
        row != bottom and Editable.get_cell(editable, coord) != :error
      end)

    if occupied, do: {:error, {:occupied, Range.to_string(target)}}, else: {:ok, book}
  end

  defp totals_destination(nil, _new_bottom, _left, _right), do: []

  defp totals_destination(_row, new_bottom, left, right),
    do: for(col <- left..right//1, do: {new_bottom, col})

  defp move_totals_cells(book, %Entry{table: %Table{totals_row_count: 0}}, _from, _to), do: book

  defp move_totals_cells(book, %Entry{table: table} = entry, from, to) do
    {_, left} = table.ref.top_left
    {_, right} = table.ref.bottom_right
    Enum.reduce(left..right//1, book, &Cells.move(&2, entry.sheet_path, {from, &1}, {to, &1}))
  end

  defp write_rows(book, %Entry{table: table} = entry, lists, first_row) do
    {_, left} = table.ref.top_left

    lists
    |> Enum.with_index(first_row)
    |> Enum.reduce(book, fn {values, row}, acc ->
      values
      |> Enum.zip(table.columns)
      |> Enum.with_index(left)
      |> Enum.reduce(acc, fn {{value, column}, col}, acc2 ->
        write_value(acc2, entry.sheet_path, {row, col}, value, column)
      end)
    end)
  end

  defp write_value(book, sheet_path, coord, nil, column) do
    case Column.calculated_formula(column) do
      nil -> book
      formula -> Cells.write(book, sheet_path, coord, {:formula, formula})
    end
  end

  defp write_value(book, sheet_path, coord, value, _column),
    do: Cells.write(book, sheet_path, coord, value)

  defp check_columns(table, names) do
    Enum.find_value(names, :ok, fn name ->
      case Validation.check_column_exists(table, name) do
        :ok -> nil
        error -> error
      end
    end)
  end

  defp check_specs(specs) do
    Enum.find_value(specs, :ok, fn {_column, spec} ->
      if valid_spec?(spec), do: nil, else: {:error, {:invalid_totals_function, spec}}
    end)
  end

  defp valid_spec?({:custom, formula}) when is_binary(formula), do: true
  defp valid_spec?({:label, text}) when is_binary(text), do: true

  defp valid_spec?(function) when is_atom(function),
    do: function == :none or Column.subtotal_code(function) != nil

  defp valid_spec?(_), do: false

  defp ensure_totals_row(book, %Entry{table: %Table{totals_row_count: 1} = table}),
    do: {:ok, table, book}

  defp ensure_totals_row(book, %Entry{table: table} = entry) do
    {_, left} = table.ref.top_left
    {bottom, right} = table.ref.bottom_right
    totals = %Range{top_left: {bottom + 1, left}, bottom_right: {bottom + 1, right}}

    case Cells.empty?(book, entry.sheet_path, totals) do
      {true, book} ->
        new_table =
          table
          |> Table.put_ref(%{table.ref | bottom_right: {bottom + 1, right}})
          |> Table.put_totals_row_count(1)

        {:ok, new_table, book}

      {false, _book} ->
        {:error, {:occupied, Range.to_string(totals)}}
    end
  end

  defp put_column_totals(table, column_name, spec) do
    {:ok, column} = Table.fetch_column(table, column_name)
    Table.put_column(table, column_name, Column.put_totals(column, spec))
  end

  defp write_totals_cell(book, sheet_path, table, {column_name, spec}) do
    {:ok, offset} = Table.column_index(table, column_name)
    {_, left} = table.ref.top_left
    coord = {Table.totals_row(table), left + offset}
    {:ok, column} = Table.fetch_column(table, column_name)

    case totals_cell_value(spec, table.name, column.name) do
      nil -> Cells.clear(book, sheet_path, coord)
      value -> Cells.write(book, sheet_path, coord, value)
    end
  end

  defp totals_cell_value({:label, text}, _table, _column), do: text
  defp totals_cell_value({:custom, formula}, _table, _column), do: {:formula, formula}
  defp totals_cell_value(:none, _table, _column), do: nil

  defp totals_cell_value(function, table, column) do
    {:formula, "SUBTOTAL(#{Column.subtotal_code(function)},#{table}[#{column}])"}
  end

  defp remove_totals_row(book, %Entry{table: %Table{totals_row_count: 0}}), do: {:ok, book}

  defp remove_totals_row(book, %Entry{table: table} = entry) do
    {_, left} = table.ref.top_left
    {bottom, right} = table.ref.bottom_right

    book = Enum.reduce(left..right//1, book, &Cells.clear(&2, entry.sheet_path, {bottom, &1}))

    new_table =
      table
      |> Table.put_totals_row_count(0)
      |> Table.put_ref(%{table.ref | bottom_right: {bottom - 1, right}})
      |> Map.update!(:columns, fn columns -> Enum.map(columns, &Column.put_totals(&1, nil)) end)

    {:ok, TableParts.put(book, entry, new_table)}
  end
end
