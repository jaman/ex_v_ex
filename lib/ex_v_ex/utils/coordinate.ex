defmodule ExVEx.Utils.Coordinate do
  @moduledoc """
  A1-style cell coordinate parsing and emission.

  Coordinates are represented internally as `{row, col}` 1-indexed tuples.
  Columns use Excel's bijective base-26: `A..Z, AA..AZ, .., XFD`.
  """

  @type row :: pos_integer()
  @type col :: pos_integer()
  @type t :: {row(), col()}

  @spec parse(String.t()) :: {:ok, t()} | :error
  def parse(binary) when is_binary(binary) do
    with {letters, digits} <- split_letters_and_digits(binary),
         {:ok, col} <- parse_column(letters),
         {:ok, row} <- parse_row(digits) do
      {:ok, {row, col}}
    end
  end

  @spec to_string(t()) :: String.t()
  def to_string({row, col}) when is_integer(row) and row > 0 and is_integer(col) and col > 0 do
    column_label(col) <> Integer.to_string(row)
  end

  @spec column_label(col()) :: String.t()
  def column_label(n) when is_integer(n) and n > 0 do
    n |> label_chars([]) |> IO.iodata_to_binary()
  end

  @spec column_number(String.t()) :: col()
  def column_number(letters) when is_binary(letters) do
    letters
    |> String.upcase()
    |> :erlang.binary_to_list()
    |> Enum.reduce(0, fn ch, acc when ch >= ?A and ch <= ?Z ->
      acc * 26 + (ch - ?A + 1)
    end)
  end

  defp label_chars(0, acc), do: acc

  defp label_chars(n, acc) do
    rem = Integer.mod(n - 1, 26)
    label_chars(div(n - 1, 26), [?A + rem | acc])
  end

  defp split_letters_and_digits(binary) do
    upper = String.upcase(binary)

    case Regex.run(~r/^([A-Z]+)([1-9][0-9]*)$/, upper) do
      [_, letters, digits] -> {letters, digits}
      _ -> :error
    end
  end

  defp parse_column(letters) do
    n = column_number(letters)
    if n >= 1, do: {:ok, n}, else: :error
  end

  defp parse_row(digits) do
    case Integer.parse(digits) do
      {n, ""} when n >= 1 -> {:ok, n}
      _ -> :error
    end
  end
end
