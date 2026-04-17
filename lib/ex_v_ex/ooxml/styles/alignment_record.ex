defmodule ExVEx.OOXML.Styles.AlignmentRecord do
  @moduledoc """
  Raw alignment attributes attached to an `<xf>` in `xl/styles.xml`. This
  is the internal OOXML-shaped representation; end users read alignment
  off the flattened `%ExVEx.Style.Alignment{}` returned by
  `ExVEx.get_style/3`.
  """

  defstruct horizontal: :general,
            vertical: :bottom,
            wrap_text: false,
            text_rotation: 0,
            indent: 0,
            shrink_to_fit: false

  @type t :: %__MODULE__{
          horizontal: atom(),
          vertical: atom(),
          wrap_text: boolean(),
          text_rotation: integer(),
          indent: non_neg_integer(),
          shrink_to_fit: boolean()
        }
end
