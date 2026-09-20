defmodule ExVEx.Tables.Cells do
  @moduledoc """
  Cell reads and writes on a worksheet part, addressed by part path,
  for the table operations. Values go through `ExVEx.CellCodec`, so
  strings are interned and dates are styled exactly as `ExVEx.put_cell/4`
  does.
  """

  alias ExVEx.CellCodec
  alias ExVEx.OOXML.Worksheet.Editable
  alias ExVEx.Utils.{Coordinate, Range}
  alias ExVEx.Workbook

  @spec read(Workbook.t(), String.t(), Coordinate.t()) :: {:ok, term(), Workbook.t()}
  def read(%Workbook{} = book, sheet_path, coord) do
    {:ok, editable, book} = Workbook.fetch_sheet_tree(book, sheet_path)

    value =
      case Editable.cell_record_at(editable, coord) do
        {:ok, cell} -> cell |> CellCodec.decode(book) |> decoded_or_nil()
        :error -> nil
      end

    {:ok, value, book}
  end

  @spec formula(Workbook.t(), String.t(), Coordinate.t()) :: {:ok, String.t() | nil, Workbook.t()}
  def formula(%Workbook{} = book, sheet_path, coord) do
    {:ok, editable, book} = Workbook.fetch_sheet_tree(book, sheet_path)

    case Editable.cell_record_at(editable, coord) do
      {:ok, %{formula: formula}} -> {:ok, formula, book}
      :error -> {:ok, nil, book}
    end
  end

  @spec write(Workbook.t(), String.t(), Coordinate.t(), term()) :: Workbook.t()
  def write(%Workbook{} = book, sheet_path, coord, value) do
    {:ok, editable, book} = Workbook.fetch_sheet_tree(book, sheet_path)
    {encoded, book} = CellCodec.encode(book, value)
    Workbook.put_sheet_tree(book, sheet_path, Editable.put_cell(editable, coord, encoded))
  end

  @doc "Writes `value` only when the cell does not already hold it."
  @spec write_if_different(Workbook.t(), String.t(), Coordinate.t(), term()) :: Workbook.t()
  def write_if_different(%Workbook{} = book, sheet_path, coord, value) do
    case read(book, sheet_path, coord) do
      {:ok, ^value, book} -> book
      {:ok, _other, book} -> write(book, sheet_path, coord, value)
    end
  end

  @spec clear(Workbook.t(), String.t(), Coordinate.t()) :: Workbook.t()
  def clear(%Workbook{} = book, sheet_path, coord) do
    {:ok, editable, book} = Workbook.fetch_sheet_tree(book, sheet_path)
    Workbook.put_sheet_tree(book, sheet_path, Editable.put_cell(editable, coord, nil))
  end

  @spec move(Workbook.t(), String.t(), Coordinate.t(), Coordinate.t()) :: Workbook.t()
  def move(%Workbook{} = book, sheet_path, from, to) do
    {:ok, editable, book} = Workbook.fetch_sheet_tree(book, sheet_path)
    Workbook.put_sheet_tree(book, sheet_path, Editable.move_cell(editable, from, to))
  end

  @doc "Whether every cell in `range` is absent from the sheet."
  @spec empty?(Workbook.t(), String.t(), Range.t()) :: {boolean(), Workbook.t()}
  def empty?(%Workbook{} = book, sheet_path, %Range{} = range) do
    {:ok, editable, book} = Workbook.fetch_sheet_tree(book, sheet_path)
    {Enum.all?(Range.cells(range), &(Editable.get_cell(editable, &1) == :error)), book}
  end

  @spec merged_ranges(Workbook.t(), String.t()) :: {[Range.t()], Workbook.t()}
  def merged_ranges(%Workbook{} = book, sheet_path) do
    {:ok, editable, book} = Workbook.fetch_sheet_tree(book, sheet_path)
    {Editable.merged_ranges(editable), book}
  end

  defp decoded_or_nil({:ok, value}), do: value
  defp decoded_or_nil({:error, _}), do: nil
end
