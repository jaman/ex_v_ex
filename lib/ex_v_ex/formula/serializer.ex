defmodule ExVEx.Formula.Serializer do
  @moduledoc """
  Turns a list of `ExVEx.Formula.Token` records back into the formula
  string. Each token carries its canonical textual form in the `:text`
  field (populated by the tokenizer) OR derived from structured fields
  after a shift operation.
  """

  alias ExVEx.Formula.{Reference, Token}

  @spec to_string([Token.t()]) :: String.t()
  def to_string(tokens), do: tokens |> Enum.map(&emit/1) |> IO.iodata_to_binary()

  # When `text` is non-nil, the token has not been mutated since tokenisation
  # — emit the original substring verbatim. Shift operations clear `text`
  # (see `ExVEx.Formula.Shift`) to signal "rebuild from structured fields."
  defp emit(%Token{text: text}) when is_binary(text), do: text

  defp emit(%Token{kind: :literal, text: nil} = t), do: t.text || ""

  defp emit(%Token{kind: :cell_ref, sheet: sheet, ref: ref}) do
    sheet_prefix(sheet) <> Reference.to_string(ref)
  end

  defp emit(%Token{kind: :range_ref, sheet: sheet, start_ref: a, end_ref: b}) do
    sheet_prefix(sheet) <> Reference.to_string(a) <> ":" <> Reference.to_string(b)
  end

  defp emit(%Token{kind: :row_range} = t) do
    sheet_prefix(t.sheet) <>
      abs_prefix(t.start_abs?) <>
      Integer.to_string(t.start_row) <>
      ":" <>
      abs_prefix(t.end_abs?) <>
      Integer.to_string(t.end_row)
  end

  defp emit(%Token{kind: :col_range} = t) do
    sheet_prefix(t.sheet) <>
      abs_prefix(t.start_abs?) <>
      column_label(t.start_col) <>
      ":" <>
      abs_prefix(t.end_abs?) <>
      column_label(t.end_col)
  end

  defp sheet_prefix(nil), do: ""

  defp sheet_prefix(sheet) when is_binary(sheet) do
    if Regex.match?(~r/^[A-Za-z_][A-Za-z0-9_.]*(?::[A-Za-z_][A-Za-z0-9_.]*)?$/, sheet) do
      sheet <> "!"
    else
      "'" <> String.replace(sheet, "'", "''") <> "'!"
    end
  end

  defp abs_prefix(true), do: "$"
  defp abs_prefix(false), do: ""

  defp column_label(n) do
    n |> label_chars([]) |> IO.iodata_to_binary()
  end

  defp label_chars(0, acc), do: acc

  defp label_chars(n, acc) do
    rem = Integer.mod(n - 1, 26)
    label_chars(div(n - 1, 26), [?A + rem | acc])
  end
end
