defmodule ExVEx.OOXML.Workbook.SheetRef do
  @moduledoc """
  A reference to a worksheet as declared in `xl/workbook.xml`.

  The `rel_id` resolves against `xl/_rels/workbook.xml.rels` to produce the
  actual package path of the sheet (typically `xl/worksheets/sheet*.xml`).
  """

  @enforce_keys [:name, :sheet_id, :rel_id]
  defstruct [:name, :sheet_id, :rel_id, state: :visible]

  @type state :: :visible | :hidden | :very_hidden
  @type t :: %__MODULE__{
          name: String.t(),
          sheet_id: pos_integer(),
          rel_id: String.t(),
          state: state()
        }
end
