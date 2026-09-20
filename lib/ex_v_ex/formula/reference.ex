defmodule ExVEx.Formula.Reference do
  @moduledoc """
  Absolute-or-relative A1-style cell references, as they appear inside
  Excel formula tokens.

  `row_abs?` / `col_abs?` indicate the presence of the `$` prefix on each
  axis. A fully-absolute reference `$A$1` has both true; a relative one
  `A1` has both false.
  """

  @enforce_keys [:row, :col, :row_abs?, :col_abs?]
  defstruct [:row, :col, :row_abs?, :col_abs?]

  @type t :: %__MODULE__{
          row: pos_integer(),
          col: pos_integer(),
          row_abs?: boolean(),
          col_abs?: boolean()
        }

  @max_row 1_048_576
  @max_col 16_384

  @spec parse(String.t()) :: {:ok, t()} | :error
  def parse(ref) when is_binary(ref) do
    case Regex.run(~r/^(\$?)([A-Za-z]+)(\$?)([1-9][0-9]*)$/, ref) do
      [_, col_abs, letters, row_abs, digits] ->
        build(String.to_integer(digits), column_number(letters), row_abs == "$", col_abs == "$")

      _ ->
        :error
    end
  end

  defp build(row, col, row_abs?, col_abs?) when row <= @max_row and col <= @max_col do
    {:ok, %__MODULE__{row: row, col: col, row_abs?: row_abs?, col_abs?: col_abs?}}
  end

  defp build(_row, _col, _row_abs?, _col_abs?), do: :error

  @spec to_string(t()) :: String.t()
  def to_string(%__MODULE__{} = r) do
    col_prefix = if r.col_abs?, do: "$", else: ""
    row_prefix = if r.row_abs?, do: "$", else: ""
    col_prefix <> column_label(r.col) <> row_prefix <> Integer.to_string(r.row)
  end

  defp column_number(letters) do
    letters
    |> String.upcase()
    |> :erlang.binary_to_list()
    |> Enum.reduce(0, fn ch, acc -> acc * 26 + (ch - ?A + 1) end)
  end

  defp column_label(n) do
    n |> label_chars([]) |> IO.iodata_to_binary()
  end

  defp label_chars(0, acc), do: acc

  defp label_chars(n, acc) do
    rem = Integer.mod(n - 1, 26)
    label_chars(div(n - 1, 26), [?A + rem | acc])
  end
end
