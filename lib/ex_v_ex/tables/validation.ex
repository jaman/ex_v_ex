defmodule ExVEx.Tables.Validation do
  @moduledoc """
  Checks the constraints Excel places on tables: name syntax and
  uniqueness, minimum size, no overlap with other tables or merged
  ranges, and unique column names.
  """

  alias ExVEx.OOXML.Table
  alias ExVEx.Tables.Cells
  alias ExVEx.Utils.{Coordinate, Range}
  alias ExVEx.Workbook
  alias ExVEx.Workbook.TableParts

  @name_pattern ~r/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/
  @r1c1_pattern ~r/^[Rr][0-9]*([Cc][0-9]*)?$/

  @spec check_name(Workbook.t(), String.t(), String.t() | nil) :: :ok | {:error, term()}
  def check_name(%Workbook{} = book, name, current \\ nil) do
    cond do
      not valid_name_syntax?(name) -> {:error, {:invalid_table_name, name}}
      name_taken?(book, name, current) -> {:error, {:duplicate_table_name, name}}
      true -> :ok
    end
  end

  @spec valid_name_syntax?(String.t()) :: boolean()
  def valid_name_syntax?(name) when is_binary(name) do
    String.length(name) <= 255 and Regex.match?(@name_pattern, name) and
      Coordinate.parse(name) == :error and not Regex.match?(@r1c1_pattern, name)
  end

  def valid_name_syntax?(_), do: false

  defp name_taken?(book, name, current) do
    wanted = String.downcase(name)
    skip = current && String.downcase(current)

    table_names = book |> TableParts.list() |> Enum.map(&String.downcase(&1.table.name))
    defined = Enum.map(book.workbook.defined_names, &String.downcase(&1.name))

    wanted != skip and (wanted in table_names or wanted in defined)
  end

  @doc "The first free `TableN` name."
  @spec default_name(Workbook.t()) :: String.t()
  def default_name(%Workbook{} = book), do: default_name(book, TableParts.next_id(book))

  defp default_name(book, n) do
    candidate = "Table#{n}"
    if name_taken?(book, candidate, nil), do: default_name(book, n + 1), else: candidate
  end

  @spec check_size(Range.t(), 0 | 1, 0 | 1) :: :ok | {:error, :range_too_small}
  def check_size(%Range{top_left: {top, _}, bottom_right: {bottom, _}}, header_rows, totals_rows) do
    if bottom - top + 1 >= header_rows + totals_rows + 1,
      do: :ok,
      else: {:error, :range_too_small}
  end

  @spec check_overlap(Workbook.t(), String.t(), String.t(), Range.t(), String.t() | nil) ::
          {:ok, Workbook.t()} | {:error, term()}
  def check_overlap(
        %Workbook{} = book,
        sheet_name,
        sheet_path,
        %Range{} = range,
        skip_table \\ nil
      ) do
    other_tables =
      book
      |> TableParts.sheet_entries(sheet_name, sheet_path)
      |> Enum.reject(
        &(skip_table && String.downcase(&1.table.name) == String.downcase(skip_table))
      )

    {merged, book} = Cells.merged_ranges(book, sheet_path)

    cond do
      entry = Enum.find(other_tables, &Range.overlaps?(&1.table.ref, range)) ->
        {:error, {:overlaps_table, entry.table.name}}

      merge = Enum.find(merged, &Range.overlaps?(&1, range)) ->
        {:error, {:overlaps_merged_range, Range.to_string(merge)}}

      true ->
        {:ok, book}
    end
  end

  @spec check_column_names([String.t()]) :: :ok | {:error, term()}
  def check_column_names(names) do
    Enum.reduce_while(names, MapSet.new(), fn name, seen ->
      key = is_binary(name) && String.downcase(name)

      cond do
        not is_binary(name) or name == "" -> {:halt, {:error, {:invalid_column_name, name}}}
        MapSet.member?(seen, key) -> {:halt, {:error, {:duplicate_column_name, name}}}
        true -> {:cont, MapSet.put(seen, key)}
      end
    end)
    |> case do
      {:error, _} = error -> error
      _seen -> :ok
    end
  end

  @spec check_column_exists(Table.t(), String.t()) ::
          :ok | {:error, {:unknown_column, String.t()}}
  def check_column_exists(%Table{} = table, name) do
    case Table.column_index(table, name) do
      {:ok, _} -> :ok
      :error -> {:error, {:unknown_column, name}}
    end
  end
end
