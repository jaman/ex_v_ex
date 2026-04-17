defmodule ExVEx.Style do
  @moduledoc """
  A resolved, flattened style for a single cell.

  Cells in OOXML reference a style by an integer index into `cellXfs`; each
  `<xf>` there points at a font / fill / border / number format by further
  indices. `ExVEx.get_style/3` walks those indirections and returns a single
  `%ExVEx.Style{}` record with the concrete sub-records inline.
  """

  alias ExVEx.Style.{Alignment, Border, Fill, Font}

  defstruct font: %Font{},
            fill: %Fill{},
            border: %Border{},
            alignment: %Alignment{},
            number_format: "General"

  @type t :: %__MODULE__{
          font: Font.t(),
          fill: Fill.t(),
          border: Border.t(),
          alignment: Alignment.t(),
          number_format: String.t()
        }
end
