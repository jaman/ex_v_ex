defmodule ExVEx.Style.Color do
  @moduledoc """
  A colour reference from the OOXML style model. Colours can be declared
  as a concrete RGB value, a reference to a theme palette entry, or a
  reference to the legacy indexed palette.
  """

  @enforce_keys [:kind]
  defstruct [:kind, :value, :tint]

  @type kind :: :rgb | :theme | :indexed | :auto
  @type t :: %__MODULE__{
          kind: kind(),
          value: String.t() | non_neg_integer() | nil,
          tint: float() | nil
        }
end
