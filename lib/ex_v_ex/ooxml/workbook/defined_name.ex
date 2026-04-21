defmodule ExVEx.OOXML.Workbook.DefinedName do
  @moduledoc """
  A single `<definedName>` record from `xl/workbook.xml`. Named ranges
  and named formulas live here.

  `scope` is `:global` for workbook-wide names or a sheet index
  (`{:sheet, sheet_id}`) for sheet-local names. The reference string is
  stored verbatim — static cell ranges (`"Sheet1!$A$1:$A$10"`) and
  dynamic formulas (`"=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)"`)
  are both supported; Excel evaluates them when the file is opened.
  """

  @enforce_keys [:name, :reference]
  defstruct [:name, :reference, scope: :global, hidden: false]

  @type scope :: :global | {:sheet, non_neg_integer()}
  @type t :: %__MODULE__{
          name: String.t(),
          reference: String.t(),
          scope: scope(),
          hidden: boolean()
        }
end
