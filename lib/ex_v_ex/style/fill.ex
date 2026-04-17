defmodule ExVEx.Style.Fill do
  @moduledoc """
  A fill record (the cell background) — a pattern type plus optional
  foreground and background colours.
  """

  alias ExVEx.Style.Color

  defstruct pattern: :none, foreground_color: nil, background_color: nil

  @type pattern ::
          :none
          | :solid
          | :gray125
          | :darkGray
          | :mediumGray
          | :lightGray
          | :darkHorizontal
          | :darkVertical
          | :darkDown
          | :darkUp
          | :darkGrid
          | :darkTrellis
          | :lightHorizontal
          | :lightVertical
          | :lightDown
          | :lightUp
          | :lightGrid
          | :lightTrellis
          | :gray0625
          | atom()

  @type t :: %__MODULE__{
          pattern: pattern(),
          foreground_color: Color.t() | nil,
          background_color: Color.t() | nil
        }
end
