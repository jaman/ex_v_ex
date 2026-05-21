defmodule ExVEx.OOXML.SqrefShift do
  @moduledoc """
  Shifts an OOXML `sqref` string — a space-separated list of cell
  references and/or rectangular ranges (e.g. `"A1:B2 D4 F1:F10"`) — under
  a `%ExVEx.Mutation.Shift{}`.

  Individual tokens that fall entirely inside a deletion span are
  dropped; partial overlaps clamp to the surviving portion. The returned
  string preserves the original space-separation order with deleted
  tokens removed.
  """

  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.Utils.{Coordinate, Range}

  @spec shift(String.t(), MutShift.t()) :: String.t()
  def shift(sqref, %MutShift{} = shift) when is_binary(sqref) do
    sqref
    |> String.split(~r/\s+/, trim: true)
    |> Enum.flat_map(&shift_token(&1, shift))
    |> Enum.join(" ")
  end

  defp shift_token(token, shift) do
    case String.contains?(token, ":") do
      true -> shift_range_token(token, shift)
      false -> shift_cell_token(token, shift)
    end
  end

  defp shift_cell_token(token, shift) do
    case Coordinate.parse(token) do
      {:ok, {row, col}} -> shift_cell(row, col, shift, token)
      :error -> [token]
    end
  end

  defp shift_cell(row, col, %MutShift{axis: :row} = shift, original) do
    case MutShift.apply_index(shift, row) do
      {:ok, new_row} -> [Coordinate.to_string({new_row, col})]
      :unchanged -> [original]
      :deleted -> []
    end
  end

  defp shift_cell(row, col, %MutShift{axis: :col} = shift, original) do
    case MutShift.apply_index(shift, col) do
      {:ok, new_col} -> [Coordinate.to_string({row, new_col})]
      :unchanged -> [original]
      :deleted -> []
    end
  end

  defp shift_range_token(token, shift) do
    case Range.parse(token) do
      {:ok, range} -> shift_range(range, shift, token)
      :error -> [token]
    end
  end

  defp shift_range(%Range{top_left: {tr, tc}, bottom_right: {br, bc}}, shift, original) do
    {start_idx, end_idx} = range_axis(shift.axis, tr, tc, br, bc)

    case {MutShift.apply_index(shift, start_idx), MutShift.apply_index(shift, end_idx)} do
      {:unchanged, :unchanged} ->
        [original]

      {:deleted, :deleted} ->
        []

      {s, e} ->
        new_start = extract_index(s, start_idx)
        new_end = extract_index(e, end_idx)
        [build_range(shift.axis, tr, tc, br, bc, new_start, new_end)]
    end
  end

  defp range_axis(:row, tr, _tc, br, _bc), do: {tr, br}
  defp range_axis(:col, _tr, tc, _br, bc), do: {tc, bc}

  defp extract_index({:ok, n}, _old), do: n
  defp extract_index(:unchanged, old), do: old
  defp extract_index(:deleted, old), do: old

  defp build_range(:row, _tr, tc, _br, bc, new_start, new_end) do
    Range.to_string(%Range{top_left: {new_start, tc}, bottom_right: {new_end, bc}})
  end

  defp build_range(:col, tr, _tc, br, _bc, new_start, new_end) do
    Range.to_string(%Range{top_left: {tr, new_start}, bottom_right: {br, new_end}})
  end
end
