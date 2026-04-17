defmodule ExVEx.Style.Border do
  @moduledoc "The four sides of a cell border as `%ExVEx.Style.Side{}` records."

  alias ExVEx.Style.Side

  defstruct top: %Side{}, bottom: %Side{}, left: %Side{}, right: %Side{}

  @type t :: %__MODULE__{
          top: Side.t(),
          bottom: Side.t(),
          left: Side.t(),
          right: Side.t()
        }
end
