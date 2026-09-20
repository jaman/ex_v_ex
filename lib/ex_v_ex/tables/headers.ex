defmodule ExVEx.Tables.Headers do
  @moduledoc """
  Keeps a table's column names and its header-row cells equal.

  A header cell that holds text supplies the column name; any other
  value is converted to text; an empty cell receives a placeholder
  `ColumnN`. Names are made unique within the table (case-insensitive)
  by appending a number.
  """

  alias ExVEx.OOXML.Table
  alias ExVEx.Tables.Cells
  alias ExVEx.Utils.Range
  alias ExVEx.Workbook

  @doc """
  Column names for a table over `range`: the explicit `columns` when
  given (validated for count), otherwise names derived from the header
  cells, or placeholders when the table has no header row.
  """
  @spec resolve(Workbook.t(), String.t(), Range.t(), [String.t()] | nil, boolean()) ::
          {:ok, [String.t()], Workbook.t()} | {:error, term()}
  def resolve(book, _sheet_path, range, columns, _header_row?) when is_list(columns) do
    if length(columns) == width(range),
      do: {:ok, columns, book},
      else: {:error, :column_count_mismatch}
  end

  def resolve(book, sheet_path, range, nil, true) do
    {names, book} = names_from_cells(book, sheet_path, range, 0)
    {:ok, names, book}
  end

  def resolve(book, _sheet_path, range, nil, false) do
    {:ok, Enum.map(1..width(range)//1, &"Column#{&1}"), book}
  end

  @doc "Names derived from the header cells of `range`, starting at column offset `from`."
  @spec names_from_cells(Workbook.t(), String.t(), Range.t(), non_neg_integer()) ::
          {[String.t()], Workbook.t()}
  def names_from_cells(book, sheet_path, %Range{top_left: {top, left}} = range, from) do
    {names, book} =
      Enum.map_reduce(from..(width(range) - 1)//1, book, fn offset, acc ->
        {:ok, value, acc} = Cells.read(acc, sheet_path, {top, left + offset})
        {name_from_value(value, offset + 1), acc}
      end)

    {names, book}
  end

  @doc "Writes each column name into its header cell where the cell differs."
  @spec write(Workbook.t(), String.t(), Table.t()) :: Workbook.t()
  def write(book, _sheet_path, %Table{header_row_count: 0}), do: book

  def write(book, sheet_path, %Table{} = table) do
    {top, left} = table.ref.top_left

    table.columns
    |> Enum.with_index()
    |> Enum.reduce(book, fn {column, offset}, acc ->
      Cells.write_if_different(acc, sheet_path, {top, left + offset}, column.name)
    end)
  end

  @doc """
  After a structural shift: adopts header-cell text as column names,
  makes the names unique, and writes back any cell that differs.
  """
  @spec reconcile(Workbook.t(), String.t(), Table.t()) :: {Table.t(), Workbook.t()}
  def reconcile(book, _sheet_path, %Table{header_row_count: 0} = table), do: {table, book}

  def reconcile(book, sheet_path, %Table{} = table) do
    {top, left} = table.ref.top_left

    {names, book} =
      table.columns
      |> Enum.with_index()
      |> Enum.map_reduce(book, fn {column, offset}, acc ->
        {:ok, value, acc} = Cells.read(acc, sheet_path, {top, left + offset})
        {adopt(value, column.name, offset + 1), acc}
      end)

    new_table = Table.put_column_names(table, unique(names))
    {new_table, write(book, sheet_path, new_table)}
  end

  @doc "Makes names unique (case-insensitive) by appending 2, 3, … to repeats."
  @spec unique([String.t()]) :: [String.t()]
  def unique(names) do
    {result, _taken} =
      Enum.map_reduce(names, MapSet.new(), fn name, taken ->
        candidate = dedupe(name, taken, 2)
        {candidate, MapSet.put(taken, String.downcase(candidate))}
      end)

    result
  end

  defp dedupe(name, taken, n) do
    if MapSet.member?(taken, String.downcase(name)) do
      dedupe_numbered(name, taken, n)
    else
      name
    end
  end

  defp dedupe_numbered(name, taken, n) do
    candidate = "#{name}#{n}"

    if MapSet.member?(taken, String.downcase(candidate)),
      do: dedupe_numbered(name, taken, n + 1),
      else: candidate
  end

  defp adopt(value, _current, _position) when is_binary(value) and value != "", do: value
  defp adopt(nil, current, _position), do: current
  defp adopt("", current, _position), do: current
  defp adopt(value, _current, position), do: name_from_value(value, position)

  defp name_from_value(value, _position) when is_binary(value) and value != "", do: value
  defp name_from_value(nil, position), do: "Column#{position}"
  defp name_from_value("", position), do: "Column#{position}"
  defp name_from_value(true, _position), do: "TRUE"
  defp name_from_value(false, _position), do: "FALSE"
  defp name_from_value(value, _position), do: to_string(value)

  defp width(%Range{top_left: {_, left}, bottom_right: {_, right}}), do: right - left + 1
end
