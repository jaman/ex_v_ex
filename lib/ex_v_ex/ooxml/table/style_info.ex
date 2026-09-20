defmodule ExVEx.OOXML.Table.StyleInfo do
  @moduledoc """
  The `<tableStyleInfo>` element of a table part: the named table style
  plus the four banding/emphasis switches.

  `name` is a built-in style name such as `"TableStyleMedium2"` or the
  name of a custom style declared in `styles.xml`. `attrs` carries any
  attribute not modelled here so it survives re-serialization.
  """

  @type t :: %__MODULE__{
          name: String.t() | nil,
          show_first_column: boolean(),
          show_last_column: boolean(),
          show_row_stripes: boolean(),
          show_column_stripes: boolean(),
          attrs: [{String.t(), String.t()}]
        }

  defstruct name: "TableStyleMedium2",
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
            attrs: []

  @modelled ["name", "showFirstColumn", "showLastColumn", "showRowStripes", "showColumnStripes"]

  @spec from_node({String.t(), list(), list()}) :: t()
  def from_node({"tableStyleInfo", attrs, _children}) do
    %__MODULE__{
      name: attr(attrs, "name"),
      show_first_column: flag(attrs, "showFirstColumn"),
      show_last_column: flag(attrs, "showLastColumn"),
      show_row_stripes: flag(attrs, "showRowStripes"),
      show_column_stripes: flag(attrs, "showColumnStripes"),
      attrs: attrs
    }
  end

  @spec to_node(t()) :: {String.t(), list(), list()}
  def to_node(%__MODULE__{} = style) do
    modelled =
      [
        {"name", style.name},
        {"showFirstColumn", bit(style.show_first_column)},
        {"showLastColumn", bit(style.show_last_column)},
        {"showRowStripes", bit(style.show_row_stripes)},
        {"showColumnStripes", bit(style.show_column_stripes)}
      ]
      |> Enum.reject(fn {_, value} -> is_nil(value) end)

    extra = Enum.reject(style.attrs, fn {key, _} -> key in @modelled end)
    {"tableStyleInfo", merge_in_order(style.attrs, modelled) ++ extra, []}
  end

  defp merge_in_order(original, modelled) do
    ordered_keys = Enum.map(original, &elem(&1, 0))
    {present, missing} = Enum.split_with(modelled, fn {key, _} -> key in ordered_keys end)

    Enum.flat_map(ordered_keys, fn key ->
      case List.keyfind(present, key, 0) do
        nil -> []
        pair -> [pair]
      end
    end) ++ missing
  end

  defp attr(attrs, key) do
    case List.keyfind(attrs, key, 0) do
      {_, value} -> value
      nil -> nil
    end
  end

  defp flag(attrs, key), do: attr(attrs, key) in ["1", "true"]

  defp bit(true), do: "1"
  defp bit(false), do: "0"
end
