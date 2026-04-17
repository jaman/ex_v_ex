defmodule ExVEx.OOXML.Worksheet.Cell do
  @moduledoc """
  The raw, untyped-value form of a cell as it lives in worksheet XML.

  Higher-level conversion (shared string lookup, number parsing, boolean
  coercion, date formatting) happens in `ExVEx.get_cell/3`, which has access
  to the workbook's shared strings table and stylesheet.
  """

  @type raw_type ::
          :number
          | :shared_string
          | :boolean
          | :inline_string
          | :formula_string
          | :error

  @type t :: %__MODULE__{
          raw_type: raw_type(),
          raw_value: String.t() | nil,
          formula: String.t() | nil,
          style_id: non_neg_integer() | nil
        }

  defstruct raw_type: :number, raw_value: nil, formula: nil, style_id: nil
end
