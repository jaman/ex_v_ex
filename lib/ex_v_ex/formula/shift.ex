defmodule ExVEx.Formula.Shift do
  @moduledoc """
  Applies a `%ExVEx.Mutation.Shift{}` to a token stream produced by
  `ExVEx.Formula.Tokenizer`, rewriting cell and range references so
  they continue to point at the same logical cells after the shift.

  Rules (matching Excel's structural insert/delete semantics):

    * A reference is only shifted when the shift applies to its sheet.
      If the shift's `:sheet` is `nil`, it applies everywhere; if set,
      only tokens whose `:sheet` matches (or is `nil`, meaning the
      current sheet) are shifted.

    * Absolute-axis markers (`$`) are preserved through the shift but
      do NOT prevent it — Excel shifts every reference at or below an
      insert point, including `$A$5`. The `$` affects copy-paste
      semantics, not structural row/column inserts.

    * References *above* the insertion point (or before the deletion
      span) are unchanged.

    * References *at or below* the insertion point are incremented by
      the delta.

    * References *inside* a deletion span are collapsed to `#REF!`
      (the token's `text` is replaced with the error string).
  """

  alias ExVEx.Formula.{Reference, Token}
  alias ExVEx.Mutation.Shift

  @spec apply([Token.t()], Shift.t(), String.t() | nil) :: [Token.t()]
  def apply(tokens, %Shift{} = shift, current_sheet \\ nil) do
    Enum.map(tokens, &shift_token(&1, shift, current_sheet))
  end

  defp shift_token(%Token{kind: :literal} = t, _shift, _cs), do: t

  defp shift_token(%Token{kind: :cell_ref} = t, shift, cs) do
    if shift_applies?(t.sheet, shift.sheet, cs) do
      case shift_reference(t.ref, shift) do
        {:ok, new_ref} -> %{t | ref: new_ref, text: nil}
        :deleted -> %{t | text: "#REF!"}
        :unchanged -> t
      end
    else
      t
    end
  end

  defp shift_token(%Token{kind: :range_ref} = t, shift, cs) do
    if shift_applies?(t.sheet, shift.sheet, cs) do
      shift_range_token(t, shift)
    else
      t
    end
  end

  defp shift_range_token(%Token{} = t, shift) do
    a = shift_reference(t.start_ref, shift)
    b = shift_reference(t.end_ref, shift)

    if a == :deleted or b == :deleted do
      %{t | text: "#REF!"}
    else
      new_a = extract_ref(a, t.start_ref)
      new_b = extract_ref(b, t.end_ref)

      if new_a == t.start_ref and new_b == t.end_ref,
        do: t,
        else: %{t | start_ref: new_a, end_ref: new_b, text: nil}
    end
  end

  defp shift_token(%Token{kind: :row_range} = t, %Shift{axis: :row} = shift, cs) do
    if shift_applies?(t.sheet, shift.sheet, cs) do
      s = shift_axis_index(shift, t.start_row, t.start_abs?)
      e = shift_axis_index(shift, t.end_row, t.end_abs?)

      cond do
        s == :deleted or e == :deleted ->
          %{t | text: "#REF!"}

        s == :unchanged and e == :unchanged ->
          t

        true ->
          new_start = extract_index(s, t.start_row)
          new_end = extract_index(e, t.end_row)
          %{t | start_row: new_start, end_row: new_end, text: nil}
      end
    else
      t
    end
  end

  defp shift_token(%Token{kind: :col_range} = t, %Shift{axis: :col} = shift, cs) do
    if shift_applies?(t.sheet, shift.sheet, cs) do
      s = shift_axis_index(shift, t.start_col, t.start_abs?)
      e = shift_axis_index(shift, t.end_col, t.end_abs?)

      cond do
        s == :deleted or e == :deleted ->
          %{t | text: "#REF!"}

        s == :unchanged and e == :unchanged ->
          t

        true ->
          new_start = extract_index(s, t.start_col)
          new_end = extract_index(e, t.end_col)
          %{t | start_col: new_start, end_col: new_end, text: nil}
      end
    else
      t
    end
  end

  # Row shift doesn't affect col-ranges and vice versa
  defp shift_token(token, _shift, _cs), do: token

  defp extract_index({:ok, n}, _old), do: n
  defp extract_index(:unchanged, old), do: old

  defp shift_applies?(nil, nil, _cs), do: true
  defp shift_applies?(nil, shift_sheet, current_sheet), do: shift_sheet == current_sheet
  defp shift_applies?(_token_sheet, nil, _cs), do: true
  defp shift_applies?(token_sheet, shift_sheet, _cs), do: token_sheet == shift_sheet

  defp shift_reference(%Reference{} = ref, %Shift{axis: :row} = shift) do
    case shift_axis_index(shift, ref.row, ref.row_abs?) do
      {:ok, new_row} -> {:ok, %{ref | row: new_row}}
      :unchanged -> :unchanged
      :deleted -> :deleted
    end
  end

  defp shift_reference(%Reference{} = ref, %Shift{axis: :col} = shift) do
    case shift_axis_index(shift, ref.col, ref.col_abs?) do
      {:ok, new_col} -> {:ok, %{ref | col: new_col}}
      :unchanged -> :unchanged
      :deleted -> :deleted
    end
  end

  defp shift_axis_index(shift, idx, _abs?), do: Shift.apply_index(shift, idx)

  defp extract_ref({:ok, new_ref}, _old), do: new_ref
  defp extract_ref(:unchanged, old), do: old
end
