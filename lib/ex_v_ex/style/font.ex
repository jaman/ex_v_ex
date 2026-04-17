defmodule ExVEx.Style.Font do
  @moduledoc """
  A font record from `xl/styles.xml` — a name, size, colour, and the usual
  decoration flags.
  """

  alias ExVEx.Style.Color

  defstruct name: nil,
            size: nil,
            bold: false,
            italic: false,
            underline: :none,
            strike: false,
            color: nil

  @type underline :: :none | :single | :double | :single_accounting | :double_accounting
  @type t :: %__MODULE__{
          name: String.t() | nil,
          size: number() | nil,
          bold: boolean(),
          italic: boolean(),
          underline: underline(),
          strike: boolean(),
          color: Color.t() | nil
        }
end
