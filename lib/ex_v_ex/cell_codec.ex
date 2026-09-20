defmodule ExVEx.CellCodec do
  @moduledoc """
  Converts between user-facing cell values and the encoded form stored
  in a worksheet's cell grid.

  `encode/2` turns a value into what `ExVEx.OOXML.Worksheet.Editable.put_cell/3`
  accepts, interning strings in the shared-string table and adding a
  date number format to the stylesheet when needed. `decode/2` resolves
  a `%ExVEx.OOXML.Worksheet.Cell{}` back into a value, reading shared
  strings and turning date-formatted serials into `Date` /
  `NaiveDateTime`.
  """

  alias ExVEx.OOXML.SharedStrings
  alias ExVEx.OOXML.Styles
  alias ExVEx.OOXML.Worksheet.Cell
  alias ExVEx.Workbook

  @type encoded ::
          ExVEx.cell_value()
          | {:shared_string, non_neg_integer()}
          | {:styled, number(), non_neg_integer()}

  @spec encode(Workbook.t(), ExVEx.cell_value() | Date.t() | NaiveDateTime.t()) ::
          {encoded(), Workbook.t()}
  def encode(%Workbook{shared_strings: %SharedStrings{}} = book, value) when is_binary(value) do
    {index, sst} = SharedStrings.intern(book.shared_strings, value)
    {{:shared_string, index}, %{book | shared_strings: sst, shared_strings_dirty: true}}
  end

  def encode(book, %Date{} = date) do
    styled_serial(book, Date.to_gregorian_days(date) - gregorian_epoch(), 14)
  end

  def encode(book, %NaiveDateTime{} = dt) do
    days = Date.to_gregorian_days(NaiveDateTime.to_date(dt)) - gregorian_epoch()
    fraction = (dt.hour * 3600 + dt.minute * 60 + dt.second) / 86_400
    styled_serial(book, days + fraction, 22)
  end

  def encode(book, value), do: {value, book}

  @spec decode(Cell.t(), Workbook.t()) :: {:ok, term()} | {:error, term()}
  def decode(%{raw_type: :shared_string, raw_value: idx_str}, %Workbook{
        shared_strings: %SharedStrings{} = sst
      }) do
    with {idx, ""} <- Integer.parse(idx_str || ""),
         {:ok, text} <- SharedStrings.get(sst, idx) do
      {:ok, text}
    else
      _ -> {:error, :invalid_shared_string_index}
    end
  end

  def decode(%{raw_type: :shared_string}, _), do: {:error, :no_shared_string_table}
  def decode(%{raw_type: :inline_string, raw_value: text}, _), do: {:ok, text || ""}
  def decode(%{raw_type: :boolean, raw_value: "1"}, _), do: {:ok, true}
  def decode(%{raw_type: :boolean, raw_value: "0"}, _), do: {:ok, false}

  def decode(%{raw_type: :number} = cell, %Workbook{} = book) do
    case parse_number(cell.raw_value) do
      {:ok, number} when is_number(number) -> maybe_as_date(number, cell, book)
      other -> other
    end
  end

  def decode(%{raw_type: :formula_string, raw_value: text}, _), do: {:ok, text || ""}
  def decode(%{raw_type: :error, raw_value: code}, _), do: {:error, {:cell_error, code}}

  defp styled_serial(book, serial, num_fmt_id) do
    styles = book.styles || %Styles{}
    {style_id, styles} = Styles.upsert_date_format(styles, num_fmt_id)
    book = %{book | styles: styles, styles_dirty: true, styles_path: styles_path(book)}
    {{:styled, serial, style_id}, book}
  end

  defp styles_path(%Workbook{styles_path: path}) when is_binary(path), do: path
  defp styles_path(_), do: "xl/styles.xml"

  defp gregorian_epoch, do: Date.to_gregorian_days(~D[1899-12-30])

  defp maybe_as_date(number, %{style_id: nil}, _), do: {:ok, number}
  defp maybe_as_date(number, _cell, %Workbook{styles: nil}), do: {:ok, number}

  defp maybe_as_date(number, %{style_id: style_id}, %Workbook{styles: styles}) do
    with {:ok, xf} <- Styles.cell_format(styles, style_id),
         true <- Styles.date_format?(styles, xf),
         {:ok, value} <- serial_to_temporal(number, Styles.format_code(styles, xf.num_fmt_id)) do
      {:ok, value}
    else
      _ -> {:ok, number}
    end
  end

  defp serial_to_temporal(number, format_code) do
    if Regex.match?(~r/[hHsS]/, format_code),
      do: serial_to_naive_datetime(number),
      else: serial_to_date(number)
  end

  defp serial_to_date(serial) when is_number(serial) and serial >= 1 do
    days = trunc(serial)
    epoch = if days < 60, do: ~D[1899-12-31], else: ~D[1899-12-30]
    {:ok, Date.add(epoch, days)}
  end

  defp serial_to_date(_), do: {:error, :serial_out_of_range}

  defp serial_to_naive_datetime(serial) when is_number(serial) and serial >= 0 do
    days = trunc(serial)
    seconds_in_day = round((serial - days) * 86_400)

    date =
      if days == 0 do
        ~D[1899-12-30]
      else
        epoch = if days < 60, do: ~D[1899-12-31], else: ~D[1899-12-30]
        Date.add(epoch, days)
      end

    {:ok, NaiveDateTime.add(NaiveDateTime.new!(date, ~T[00:00:00]), seconds_in_day, :second)}
  end

  defp serial_to_naive_datetime(_), do: {:error, :serial_out_of_range}

  defp parse_number(nil), do: {:ok, nil}

  defp parse_number(raw) do
    if String.contains?(raw, [".", "e", "E"]) do
      case Float.parse(raw) do
        {f, ""} -> {:ok, f}
        _ -> {:error, {:bad_number, raw}}
      end
    else
      case Integer.parse(raw) do
        {n, ""} -> {:ok, n}
        _ -> {:error, {:bad_number, raw}}
      end
    end
  end
end
