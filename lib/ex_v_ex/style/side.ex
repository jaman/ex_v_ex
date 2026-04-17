defmodule ExVEx.Style.Side do
  @moduledoc "One side of a cell border: a line style and a colour."

  alias ExVEx.Style.Color

  defstruct style: :none, color: nil

  @type style ::
          :none
          | :thin
          | :medium
          | :dashed
          | :dotted
          | :thick
          | :double
          | :hair
          | :mediumDashed
          | :dashDot
          | :mediumDashDot
          | :dashDotDot
          | :mediumDashDotDot
          | :slantDashDot
          | atom()

  @type t :: %__MODULE__{style: style(), color: Color.t() | nil}
end
