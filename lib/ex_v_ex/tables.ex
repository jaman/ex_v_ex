defmodule ExVEx.Tables do
  @moduledoc """
  Table operations behind `ExVEx.add_table/4`, `ExVEx.remove_table/2`,
  `ExVEx.rename_table/3`, `ExVEx.rename_table_column/4`,
  `ExVEx.resize_table/3`, `ExVEx.put_table_style/3`, and the table
  listing functions. Row-level operations live in `ExVEx.Tables.Rows`.

  Every mutating function returns `{:ok, workbook}` or `{:error, reason}`
  and touches only the table part, its header cells, and the formulas
  that name the table.
  """

  alias ExVEx.Formula.StructuredReference
  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.OOXML.Table
  alias ExVEx.OOXML.Table.StyleInfo
  alias ExVEx.Tables.{Cells, Headers, Validation}
  alias ExVEx.Utils.Range
  alias ExVEx.Workbook
  alias ExVEx.Workbook.{FormulaRewrite, TableParts}
  alias ExVEx.Workbook.TableParts.Entry

  @type result :: {:ok, Workbook.t()} | {:error, term()}

  @spec list(Workbook.t()) :: [ExVEx.Table.t()]
  def list(%Workbook{} = book),
    do: book |> TableParts.list() |> Enum.map(&ExVEx.Table.from_entry/1)

  @spec list(Workbook.t(), String.t()) :: {:ok, [ExVEx.Table.t()]} | {:error, :unknown_sheet}
  def list(%Workbook{} = book, sheet) do
    with {:ok, sheet_path} <- sheet_path(book, sheet) do
      {:ok,
       book |> TableParts.sheet_entries(sheet, sheet_path) |> Enum.map(&ExVEx.Table.from_entry/1)}
    end
  end

  @spec fetch(Workbook.t(), String.t()) :: {:ok, ExVEx.Table.t()} | {:error, :unknown_table}
  def fetch(%Workbook{} = book, name) do
    with {:ok, entry} <- fetch_entry(book, name), do: {:ok, ExVEx.Table.from_entry(entry)}
  end

  @spec create(Workbook.t(), String.t(), String.t(), keyword()) :: result()
  def create(%Workbook{} = book, sheet, ref, opts) do
    header_row? = Keyword.get(opts, :header_row, true)
    header_rows = if header_row?, do: 1, else: 0
    name = Keyword.get(opts, :name) || Validation.default_name(book)

    with {:ok, sheet_path} <- sheet_path(book, sheet),
         {:ok, range} <- parse_range(ref),
         :ok <- Validation.check_size(range, header_rows, 0),
         :ok <- Validation.check_name(book, name),
         {:ok, book} <- Validation.check_overlap(book, sheet, sheet_path, range),
         {:ok, names, book} <-
           Headers.resolve(book, sheet_path, range, opts[:columns], header_row?),
         :ok <- Validation.check_column_names(names) do
      table =
        Table.new(
          id: TableParts.next_id(book),
          name: name,
          ref: range,
          columns: Headers.unique(names),
          style: style_from_opts(%StyleInfo{}, opts),
          header_row: header_row?
        )

      book = Headers.write(book, sheet_path, table)

      with {:ok, _entry, book} <- TableParts.add(book, sheet, sheet_path, table) do
        {:ok, %{book | calc_dirty: true}}
      end
    end
  end

  @spec remove(Workbook.t(), String.t()) :: result()
  def remove(%Workbook{} = book, name) do
    with {:ok, entry} <- fetch_entry(book, name) do
      book = FormulaRewrite.apply(book, &convert_references(&1, &2, entry))
      {:ok, TableParts.remove(book, entry)}
    end
  end

  @spec rename(Workbook.t(), String.t(), String.t()) :: result()
  def rename(%Workbook{} = book, old, new) do
    with {:ok, entry} <- fetch_entry(book, old),
         :ok <- Validation.check_name(book, new, entry.table.name) do
      book = TableParts.put(book, entry, Table.rename(entry.table, new))

      {:ok,
       FormulaRewrite.apply(book, fn tokens, _ ->
         StructuredReference.rename_table(tokens, old, new)
       end)}
    end
  end

  @spec rename_column(Workbook.t(), String.t(), String.t(), String.t()) :: result()
  def rename_column(%Workbook{} = book, name, old, new) do
    with {:ok, entry} <- fetch_entry(book, name),
         :ok <- Validation.check_column_exists(entry.table, old),
         :ok <- check_new_column_name(entry.table, old, new) do
      table = Table.rename_column(entry.table, old, new)

      book =
        book
        |> TableParts.put(entry, table)
        |> Headers.write(entry.sheet_path, table)
        |> FormulaRewrite.apply(fn tokens, context ->
          inside? = inside_table?(context, entry)
          StructuredReference.rename_column(tokens, entry.table.name, old, new, inside: inside?)
        end)

      {:ok, book}
    end
  end

  @spec resize(Workbook.t(), String.t(), String.t()) :: result()
  def resize(%Workbook{} = book, name, ref) do
    with {:ok, entry} <- fetch_entry(book, name),
         {:ok, range} <- parse_range(ref),
         :ok <- check_header_row_stays(entry.table, range),
         :ok <-
           Validation.check_size(
             range,
             entry.table.header_row_count,
             entry.table.totals_row_count
           ),
         {:ok, book} <-
           Validation.check_overlap(book, entry.sheet, entry.sheet_path, range, entry.table.name),
         {:ok, book} <- move_totals_row(book, entry, range) do
      {names, book} = resized_column_names(book, entry, range)
      table = entry.table |> Table.put_ref(range) |> Table.put_column_names(Headers.unique(names))

      book =
        book
        |> TableParts.put(entry, table)
        |> Headers.write(entry.sheet_path, table)

      {:ok, book}
    end
  end

  @spec put_style(Workbook.t(), String.t(), keyword()) :: result()
  def put_style(%Workbook{} = book, name, opts) do
    with {:ok, entry} <- fetch_entry(book, name) do
      style = style_from_opts(entry.table.style || %StyleInfo{}, opts)
      {:ok, TableParts.put(book, entry, Table.put_style(entry.table, style))}
    end
  end

  @doc """
  Cascades a structural shift on `sheet` into the tables of that sheet:
  parts are rewritten or removed, header cells are brought back in line
  with the column names, and totals-row cells left behind by a dropped
  totals row are cleared.
  """
  @spec shift(Workbook.t(), String.t(), String.t(), MutShift.t()) :: Workbook.t()
  def shift(%Workbook{} = book, sheet, sheet_path, %MutShift{} = mut_shift) do
    {book, survivors} = TableParts.shift(book, sheet, sheet_path, mut_shift)

    Enum.reduce(survivors, book, fn {before, after_shift}, acc ->
      acc = clear_orphaned_totals(acc, before, after_shift, mut_shift)
      {table, acc} = Headers.reconcile(acc, sheet_path, after_shift.table)
      if table == after_shift.table, do: acc, else: TableParts.put(acc, after_shift, table)
    end)
  end

  @doc false
  @spec fetch_entry(Workbook.t(), String.t()) :: {:ok, Entry.t()} | {:error, :unknown_table}
  def fetch_entry(%Workbook{} = book, name) when is_binary(name) do
    case TableParts.fetch(book, name) do
      {:ok, entry} -> {:ok, entry}
      :error -> {:error, :unknown_table}
    end
  end

  def fetch_entry(_book, _name), do: {:error, :unknown_table}

  @doc false
  @spec sheet_path(Workbook.t(), String.t()) :: {:ok, String.t()} | {:error, :unknown_sheet}
  def sheet_path(book, sheet) do
    case Workbook.sheet_path(book, sheet) do
      {:ok, path} -> {:ok, path}
      :error -> {:error, :unknown_sheet}
    end
  end

  @doc false
  @spec parse_range(String.t()) :: {:ok, Range.t()} | {:error, :invalid_range}
  def parse_range(ref) when is_binary(ref) do
    case Range.parse(ref) do
      {:ok, range} -> {:ok, range}
      :error -> {:error, :invalid_range}
    end
  end

  def parse_range(_), do: {:error, :invalid_range}

  @doc false
  @spec geometry(Table.t()) :: map()
  def geometry(%Table{} = table) do
    {top, left} = table.ref.top_left
    {bottom, right} = table.ref.bottom_right

    %{
      top: top,
      bottom: bottom,
      left: left,
      right: right,
      header_row: Table.header_row(table),
      totals_row: Table.totals_row(table),
      data_rows: Table.data_rows(table),
      columns: Table.column_names(table)
    }
  end

  defp convert_references(tokens, context, %Entry{} = entry) do
    geometry = geometry(entry.table)
    prefix = if context.sheet == entry.sheet, do: "", else: sheet_prefix(entry.sheet)
    formula_row = context.coord && elem(context.coord, 0)

    Enum.map(tokens, fn token ->
      if references_table?(token, entry, context),
        do: StructuredReference.to_range(token, geometry, prefix, formula_row),
        else: token
    end)
  end

  defp references_table?(%{kind: :structured_ref, table: nil}, entry, context),
    do: inside_table?(context, entry)

  defp references_table?(%{kind: :structured_ref, table: table}, entry, _context) do
    String.downcase(table) == String.downcase(entry.table.name)
  end

  defp references_table?(_token, _entry, _context), do: false

  defp inside_table?(%{sheet: sheet, coord: coord}, %Entry{} = entry) when is_tuple(coord) do
    sheet == entry.sheet and Range.contains?(entry.table.ref, coord)
  end

  defp inside_table?(_context, _entry), do: false

  defp sheet_prefix(sheet) do
    if Regex.match?(~r/^[A-Za-z_][A-Za-z0-9_.]*$/, sheet),
      do: sheet <> "!",
      else: "'" <> String.replace(sheet, "'", "''") <> "'!"
  end

  defp check_new_column_name(table, old, new) do
    cond do
      not is_binary(new) or new == "" -> {:error, {:invalid_column_name, new}}
      String.downcase(new) == String.downcase(old) -> :ok
      match?({:ok, _}, Table.column_index(table, new)) -> {:error, {:duplicate_column_name, new}}
      true -> :ok
    end
  end

  defp check_header_row_stays(%Table{ref: %Range{top_left: {top, _}}}, %Range{top_left: {top, _}}),
       do: :ok

  defp check_header_row_stays(_table, _range), do: {:error, :header_row_must_stay}

  defp move_totals_row(book, %Entry{table: %Table{totals_row_count: 0}}, _range), do: {:ok, book}

  defp move_totals_row(book, %Entry{table: table} = entry, %Range{} = range) do
    {old_bottom, left} = table.ref.bottom_right
    {new_bottom, _} = range.bottom_right

    if old_bottom == new_bottom do
      {:ok, book}
    else
      width = elem(range.bottom_right, 1) - left
      destination = %Range{top_left: {new_bottom, left}, bottom_right: {new_bottom, left + width}}

      case Cells.empty?(book, entry.sheet_path, destination) do
        {true, book} ->
          {:ok,
           Enum.reduce(0..width//1, book, fn offset, acc ->
             Cells.move(
               acc,
               entry.sheet_path,
               {old_bottom, left + offset},
               {new_bottom, left + offset}
             )
           end)}

        {false, _book} ->
          {:error, {:occupied, Range.to_string(destination)}}
      end
    end
  end

  defp resized_column_names(book, %Entry{table: table} = entry, %Range{} = range) do
    existing = Table.column_names(table)
    new_width = elem(range.bottom_right, 1) - elem(range.top_left, 1) + 1

    cond do
      new_width <= length(existing) ->
        {Enum.take(existing, new_width), book}

      table.header_row_count == 0 ->
        {existing ++ Enum.map((length(existing) + 1)..new_width//1, &"Column#{&1}"), book}

      true ->
        {added, book} = Headers.names_from_cells(book, entry.sheet_path, range, length(existing))
        {existing ++ added, book}
    end
  end

  defp clear_orphaned_totals(
         book,
         %Entry{table: before},
         %Entry{table: after_shift} = entry,
         mut_shift
       ) do
    if before.totals_row_count == 1 and after_shift.totals_row_count == 0 do
      {old_bottom, left} = before.ref.bottom_right
      {_, right} = before.ref.bottom_right

      case MutShift.apply_index(mut_shift, old_bottom) do
        :deleted -> book
        :unchanged -> clear_row(book, entry.sheet_path, old_bottom, left, right)
        {:ok, row} -> clear_row(book, entry.sheet_path, row, left, right)
      end
    else
      book
    end
  end

  defp clear_row(book, sheet_path, row, left, right) do
    Enum.reduce(left..right//1, book, &Cells.clear(&2, sheet_path, {row, &1}))
  end

  defp style_from_opts(%StyleInfo{} = base, opts) do
    if Keyword.has_key?(opts, :style) and is_nil(opts[:style]) do
      nil
    else
      %StyleInfo{
        base
        | name: Keyword.get(opts, :style, base.name),
          show_first_column: Keyword.get(opts, :show_first_column, base.show_first_column),
          show_last_column: Keyword.get(opts, :show_last_column, base.show_last_column),
          show_row_stripes: Keyword.get(opts, :show_row_stripes, base.show_row_stripes),
          show_column_stripes: Keyword.get(opts, :show_column_stripes, base.show_column_stripes)
      }
    end
  end
end
