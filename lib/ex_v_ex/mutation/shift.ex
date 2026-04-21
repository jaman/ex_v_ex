defmodule ExVEx.Mutation.Shift do
  @moduledoc """
  A row/column insertion or deletion expressed as a single data record.

  `axis` is `:row` or `:col`. `at` is the 1-indexed position *at which*
  the insertion happens (everything at that position and below/right
  shifts); for deletion, `at` is the first position removed. `count` is
  how many rows/columns are affected. `delta` is `+count` for insertion
  and `-count` for deletion, pre-computed for convenience.

  `sheet` identifies which sheet the shift applies to; when `nil` the
  shift applies across the whole workbook (used while walking formulas
  whose ranges don't specify a sheet).
  """

  @enforce_keys [:axis, :at, :count, :delta, :sheet]
  defstruct [:axis, :at, :count, :delta, :sheet]

  @type axis :: :row | :col
  @type t :: %__MODULE__{
          axis: axis(),
          at: pos_integer(),
          count: pos_integer(),
          delta: integer(),
          sheet: String.t() | nil
        }

  @spec insert(axis(), pos_integer(), pos_integer(), String.t() | nil) :: t()
  def insert(axis, at, count, sheet) when count >= 1 do
    %__MODULE__{axis: axis, at: at, count: count, delta: count, sheet: sheet}
  end

  @spec delete(axis(), pos_integer(), pos_integer(), String.t() | nil) :: t()
  def delete(axis, at, count, sheet) when count >= 1 do
    %__MODULE__{axis: axis, at: at, count: count, delta: -count, sheet: sheet}
  end

  @doc """
  Applies the shift to a single 1-indexed coordinate on the axis, if it
  is affected. Returns `{:ok, new_index}`, `:unchanged`, or `:deleted`
  if the position falls inside a deletion span.
  """
  @spec apply_index(t(), pos_integer()) ::
          {:ok, pos_integer()} | :unchanged | :deleted
  def apply_index(%__MODULE__{at: at}, idx) when idx < at, do: :unchanged

  def apply_index(%__MODULE__{delta: d}, idx) when d > 0, do: {:ok, idx + d}

  def apply_index(%__MODULE__{delta: d, at: at, count: count}, idx) do
    if idx >= at and idx < at + count do
      :deleted
    else
      {:ok, idx + d}
    end
  end
end
