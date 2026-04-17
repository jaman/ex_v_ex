defmodule ExVEx.Packaging.ContentTypes.Default do
  @moduledoc """
  A default content type mapping: every part whose path ends in `extension`
  has `content_type` unless overridden.
  """

  @enforce_keys [:extension, :content_type]
  defstruct [:extension, :content_type]

  @type t :: %__MODULE__{extension: String.t(), content_type: String.t()}
end
