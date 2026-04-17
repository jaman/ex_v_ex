defmodule ExVEx.Packaging.ContentTypes.Override do
  @moduledoc """
  A per-part content type override. `part_name` is an absolute path with a
  leading `/` (e.g. `"/xl/workbook.xml"`) per the OPC spec.
  """

  @enforce_keys [:part_name, :content_type]
  defstruct [:part_name, :content_type]

  @type t :: %__MODULE__{part_name: String.t(), content_type: String.t()}
end
