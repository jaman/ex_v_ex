defmodule ExVEx.Utils.Range do
  @moduledoc """
  Rectangular cell ranges (`"A1:B2"` in Excel's notation).

  Internally a range is two coordinates — the top-left and bottom-right
  corners, normalised so that `top_left` is always the smaller row/column.
  Swapped corners (`"B2:A1"`) are accepted and normalised.
  """

  alias ExVEx.Utils.Coordinate

  @enforce_keys [:top_left, :bottom_right]
  defstruct [:top_left, :bottom_right]

  @type t :: %__MODULE__{top_left: Coordinate.t(), bottom_right: Coordinate.t()}

  @spec parse(String.t()) :: {:ok, t()} | :error
  def parse(binary) when is_binary(binary) do
    with [left_ref, right_ref] <- String.split(binary, ":", parts: 2),
         {:ok, left} <- Coordinate.parse(left_ref),
         {:ok, right} <- Coordinate.parse(right_ref) do
      {:ok, normalize(left, right)}
    else
      _ -> :error
    end
  end

  @spec to_string(t()) :: String.t()
  def to_string(%__MODULE__{top_left: tl, bottom_right: br}) do
    Coordinate.to_string(tl) <> ":" <> Coordinate.to_string(br)
  end

  @spec anchor(t()) :: Coordinate.t()
  def anchor(%__MODULE__{top_left: tl}), do: tl

  @spec overlaps?(t(), t()) :: boolean()
  def overlaps?(%__MODULE__{} = a, %__MODULE__{} = b) do
    {a_top, a_left} = a.top_left
    {a_bottom, a_right} = a.bottom_right
    {b_top, b_left} = b.top_left
    {b_bottom, b_right} = b.bottom_right

    a_left <= b_right and b_left <= a_right and
      a_top <= b_bottom and b_top <= a_bottom
  end

  @spec cells(t()) :: [Coordinate.t()]
  def cells(%__MODULE__{top_left: {top, left}, bottom_right: {bottom, right}}) do
    for row <- top..bottom, col <- left..right, do: {row, col}
  end

  @spec contains?(t(), Coordinate.t()) :: boolean()
  def contains?(%__MODULE__{top_left: {top, left}, bottom_right: {bottom, right}}, {row, col}) do
    row >= top and row <= bottom and col >= left and col <= right
  end

  defp normalize({r1, c1}, {r2, c2}) do
    %__MODULE__{
      top_left: {min(r1, r2), min(c1, c2)},
      bottom_right: {max(r1, r2), max(c1, c2)}
    }
  end
end
