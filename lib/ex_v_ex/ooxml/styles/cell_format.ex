defmodule ExVEx.OOXML.Styles.CellFormat do
  @moduledoc """
  A single `<xf>` record from `cellXfs`. Cells reference these by index via
  the `s` attribute (`<c r="A1" s="3"/>` → `cell_formats[3]`).
  """

  defstruct num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0, xf_id: 0

  @type t :: %__MODULE__{
          num_fmt_id: non_neg_integer(),
          font_id: non_neg_integer(),
          fill_id: non_neg_integer(),
          border_id: non_neg_integer(),
          xf_id: non_neg_integer()
        }
end
