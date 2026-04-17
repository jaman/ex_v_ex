defmodule ExVEx.OOXML.Styles.NumFmt do
  @moduledoc """
  A custom number format declared in `xl/styles.xml`. IDs 0..163 are
  reserved for built-in formats defined by the OOXML spec; IDs 164 and
  above are custom formats authored by the workbook.
  """

  @enforce_keys [:id, :format_code]
  defstruct [:id, :format_code]

  @type t :: %__MODULE__{id: non_neg_integer(), format_code: String.t()}
end
