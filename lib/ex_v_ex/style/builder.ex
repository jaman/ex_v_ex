defmodule ExVEx.Style.Builder do
  @moduledoc """
  Translates the keyword-style options accepted by `ExVEx.put_style/4`
  into updated `%Font{}`, `%Fill{}`, `%Border{}`, `%Alignment{}` records,
  merging with the cell's existing style.

  Supported options:

  Font: `:bold`, `:italic`, `:strike`, `:underline`, `:font_size`,
  `:font_name`, `:color` (ARGB `"FFRRGGBB"` or `"RRGGBB"` auto-prefixed,
  or `{:theme, n}`, `{:indexed, n}`, or `:auto`).

  Fill: `:background` (`"RRGGBB"` shorthand → solid pattern, or `:none`),
  `:fill_pattern` + `:fill_fg` + `:fill_bg` for full control.

  Border: `:border` (applies to all sides: atom like `:thin` or keyword
  `[top: :thin, bottom: :medium]`), `:border_color`.

  Alignment: `:align` (`:left` | `:center` | `:right` | `:fill` |
  `:justify` | `:general`), `:valign` (`:top` | `:center` | `:bottom` |
  `:justify`), `:wrap_text`, `:indent`, `:text_rotation`.

  Number format: `:number_format` (`"yyyy-mm-dd"`, `"#,##0.00"`, `"0%"`,
  etc.).
  """

  alias ExVEx.OOXML.Styles.AlignmentRecord
  alias ExVEx.Style.{Border, Color, Fill, Font, Side}

  @type opts :: keyword()

  @spec apply_to_font(Font.t(), opts()) :: Font.t()
  def apply_to_font(%Font{} = font, opts) do
    font
    |> maybe_set(:bold, opts[:bold])
    |> maybe_set(:italic, opts[:italic])
    |> maybe_set(:strike, opts[:strike])
    |> maybe_set_underline(opts[:underline])
    |> maybe_set(:size, opts[:font_size])
    |> maybe_set(:name, opts[:font_name])
    |> maybe_set_color(opts[:color])
  end

  @spec apply_to_fill(Fill.t(), opts()) :: Fill.t()
  def apply_to_fill(%Fill{} = fill, opts) do
    fill
    |> apply_background(opts[:background])
    |> maybe_set(:pattern, opts[:fill_pattern])
    |> maybe_set_fill_color(:foreground_color, opts[:fill_fg])
    |> maybe_set_fill_color(:background_color, opts[:fill_bg])
  end

  @spec apply_to_border(Border.t(), opts()) :: Border.t()
  def apply_to_border(%Border{} = border, opts) do
    color = parse_color(opts[:border_color])

    border
    |> apply_all_sides(opts[:border], color)
    |> apply_side(:top, opts[:border_top], color)
    |> apply_side(:bottom, opts[:border_bottom], color)
    |> apply_side(:left, opts[:border_left], color)
    |> apply_side(:right, opts[:border_right], color)
  end

  @spec apply_to_alignment(AlignmentRecord.t() | nil, opts()) :: AlignmentRecord.t() | nil
  def apply_to_alignment(existing, opts) do
    starting = existing || %AlignmentRecord{}

    updated =
      starting
      |> maybe_set(:horizontal, opts[:align])
      |> maybe_set(:vertical, opts[:valign])
      |> maybe_set(:wrap_text, opts[:wrap_text])
      |> maybe_set(:indent, opts[:indent])
      |> maybe_set(:text_rotation, opts[:text_rotation])
      |> maybe_set(:shrink_to_fit, opts[:shrink_to_fit])

    if updated == %AlignmentRecord{}, do: nil, else: updated
  end

  defp maybe_set(struct, _field, nil), do: struct
  defp maybe_set(struct, field, value), do: Map.put(struct, field, value)

  defp maybe_set_underline(font, nil), do: font
  defp maybe_set_underline(font, value), do: %{font | underline: value}

  defp maybe_set_color(font, nil), do: font
  defp maybe_set_color(font, value), do: %{font | color: parse_color(value)}

  defp apply_background(%Fill{} = fill, nil), do: fill

  defp apply_background(%Fill{} = fill, :none),
    do: %{fill | pattern: :none, foreground_color: nil, background_color: nil}

  defp apply_background(%Fill{} = fill, color_spec) do
    color = parse_color(color_spec)
    %{fill | pattern: :solid, foreground_color: color, background_color: color}
  end

  defp maybe_set_fill_color(fill, _field, nil), do: fill
  defp maybe_set_fill_color(fill, field, value), do: Map.put(fill, field, parse_color(value))

  defp apply_all_sides(%Border{} = border, nil, _color), do: border

  defp apply_all_sides(%Border{} = _border, style, color) do
    side = %Side{style: style, color: color}
    %Border{top: side, bottom: side, left: side, right: side}
  end

  defp apply_side(border, _which, nil, _color), do: border

  defp apply_side(border, which, style, color) do
    Map.put(border, which, %Side{style: style, color: color})
  end

  defp parse_color(nil), do: nil
  defp parse_color(%Color{} = c), do: c
  defp parse_color(:auto), do: %Color{kind: :auto}

  defp parse_color({:theme, n}) when is_integer(n),
    do: %Color{kind: :theme, value: n}

  defp parse_color({:theme, n, tint}) when is_integer(n),
    do: %Color{kind: :theme, value: n, tint: tint}

  defp parse_color({:indexed, n}) when is_integer(n),
    do: %Color{kind: :indexed, value: n}

  defp parse_color({:rgb, rgb}) when is_binary(rgb), do: parse_color(rgb)

  defp parse_color(rgb) when is_binary(rgb) do
    value =
      case byte_size(rgb) do
        6 -> "FF" <> String.upcase(rgb)
        8 -> String.upcase(rgb)
        _ -> rgb
      end

    %Color{kind: :rgb, value: value}
  end
end
