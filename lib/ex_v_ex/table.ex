defmodule ExVEx.Table do
  @moduledoc """
  A read-only description of an Excel table (a "ListObject") as returned
  by `ExVEx.table/2` and `ExVEx.tables/1,2`.

  Ranges are A1 strings on `sheet`. `data_range` and `totals_range` are
  `nil` when the table has no data rows or no totals row. `columns` lists
  the column names in left-to-right order; they equal the header cells.
  """

  alias ExVEx.OOXML.Table, as: TablePart
  alias ExVEx.OOXML.Table.StyleInfo
  alias ExVEx.Utils.Range
  alias ExVEx.Workbook.TableParts.Entry

  @type t :: %__MODULE__{
          name: String.t(),
          sheet: String.t(),
          ref: String.t(),
          header_range: String.t() | nil,
          data_range: String.t() | nil,
          totals_range: String.t() | nil,
          columns: [String.t()],
          style: String.t() | nil,
          show_first_column: boolean(),
          show_last_column: boolean(),
          show_row_stripes: boolean(),
          show_column_stripes: boolean()
        }

  @enforce_keys [:name, :sheet, :ref, :columns]
  defstruct [
    :name,
    :sheet,
    :ref,
    :header_range,
    :data_range,
    :totals_range,
    :columns,
    :style,
    show_first_column: false,
    show_last_column: false,
    show_row_stripes: false,
    show_column_stripes: false
  ]

  @doc false
  @spec from_entry(Entry.t()) :: t()
  def from_entry(%Entry{sheet: sheet, table: %TablePart{} = table}) do
    {_, left} = table.ref.top_left
    {_, right} = table.ref.bottom_right

    %__MODULE__{
      name: table.name,
      sheet: sheet,
      ref: Range.to_string(table.ref),
      header_range: row_range(TablePart.header_row(table), left, right),
      data_range: rows_range(TablePart.data_rows(table), left, right),
      totals_range: row_range(TablePart.totals_row(table), left, right),
      columns: TablePart.column_names(table)
    }
    |> struct!(style_fields(table.style))
  end

  defp style_fields(nil), do: []

  defp style_fields(%StyleInfo{} = style) do
    [
      style: style.name,
      show_first_column: style.show_first_column,
      show_last_column: style.show_last_column,
      show_row_stripes: style.show_row_stripes,
      show_column_stripes: style.show_column_stripes
    ]
  end

  defp row_range(nil, _left, _right), do: nil
  defp row_range(row, left, right), do: rows_range({row, row}, left, right)

  defp rows_range(nil, _left, _right), do: nil

  defp rows_range({first, last}, left, right) do
    Range.to_string(%Range{top_left: {first, left}, bottom_right: {last, right}})
  end
end
