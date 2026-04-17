defmodule ExVEx.Style.Alignment do
  @moduledoc "Alignment and wrapping options declared on a `<xf>` record."

  defstruct horizontal: :general,
            vertical: :bottom,
            wrap_text: false,
            text_rotation: 0,
            indent: 0,
            shrink_to_fit: false

  @type horizontal ::
          :general
          | :left
          | :center
          | :right
          | :fill
          | :justify
          | :center_continuous
          | :distributed
  @type vertical :: :top | :center | :bottom | :justify | :distributed

  @type t :: %__MODULE__{
          horizontal: horizontal(),
          vertical: vertical(),
          wrap_text: boolean(),
          text_rotation: integer(),
          indent: non_neg_integer(),
          shrink_to_fit: boolean()
        }
end
