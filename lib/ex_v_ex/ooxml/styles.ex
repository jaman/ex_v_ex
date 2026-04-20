defmodule ExVEx.OOXML.Styles do
  @moduledoc """
  Parser for `xl/styles.xml`.

  The OOXML style model is an indirection graph: a cell carries a style
  index `s="N"` into `cellXfs`; each `<xf>` there points at a `numFmt`,
  `font`, `fill`, and `border` by further indices.

  ExVEx parses the indirection arrays (`numFmts`, `fonts`, `fills`,
  `borders`, `cellXfs`) and exposes `resolve/2` to flatten an `<xf>`
  index into a user-facing `%ExVEx.Style{}` record.

  Built-in number format IDs (0..163) are not listed in `numFmts` at all;
  workbook-specific custom formats start at ID 164.
  """

  alias ExVEx.OOXML.Styles.{AlignmentRecord, CellFormat, NumFmt}
  alias ExVEx.Style
  alias ExVEx.Style.{Alignment, Border, Color, Fill, Font, Side}

  @type t :: %__MODULE__{
          num_fmts: [NumFmt.t()],
          fonts: [Font.t()],
          fills: [Fill.t()],
          borders: [Border.t()],
          cell_formats: [CellFormat.t()]
        }

  defstruct num_fmts: [], fonts: [], fills: [], borders: [], cell_formats: []

  @builtin_date_ids MapSet.new([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47])

  @builtin_format_codes %{
    0 => "General",
    1 => "0",
    2 => "0.00",
    3 => "#,##0",
    4 => "#,##0.00",
    9 => "0%",
    10 => "0.00%",
    11 => "0.00E+00",
    12 => "# ?/?",
    13 => "# ??/??",
    14 => "m/d/yyyy",
    15 => "d-mmm-yy",
    16 => "d-mmm",
    17 => "mmm-yy",
    18 => "h:mm AM/PM",
    19 => "h:mm:ss AM/PM",
    20 => "h:mm",
    21 => "h:mm:ss",
    22 => "m/d/yyyy h:mm",
    37 => "#,##0 ;(#,##0)",
    38 => "#,##0 ;[Red](#,##0)",
    39 => "#,##0.00;(#,##0.00)",
    40 => "#,##0.00;[Red](#,##0.00)",
    45 => "mm:ss",
    46 => "[h]:mm:ss",
    47 => "mm:ss.0",
    48 => "##0.0E+0",
    49 => "@"
  }

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"styleSheet", _attrs, children}} ->
        {:ok,
         %__MODULE__{
           num_fmts: collect_num_fmts(children),
           fonts: collect_fonts(children),
           fills: collect_fills(children),
           borders: collect_borders(children),
           cell_formats: collect_cell_formats(children)
         }}

      {:ok, _other} ->
        {:error, :not_a_stylesheet}

      {:error, reason} ->
        {:error, reason}
    end
  end

  @doc """
  Re-serialises the stylesheet into its original XML container, rewriting
  only the `<cellXfs>` section. All other elements (fonts, fills, borders,
  cellStyles, dxfs, tableStyles, etc.) are preserved at the element level
  from `original_xml`.
  """
  @spec serialize_into(t(), binary()) :: binary()
  def serialize_into(%__MODULE__{cell_formats: xfs}, original_xml) when is_binary(original_xml) do
    case Saxy.SimpleForm.parse_string(original_xml) do
      {:ok, {"styleSheet", attrs, children}} ->
        new_children = replace_cell_xfs(children, xfs)
        tree = {"styleSheet", attrs, new_children}
        Saxy.encode!(tree, version: "1.0", encoding: "UTF-8", standalone: true)

      _ ->
        original_xml
    end
  end

  defp replace_cell_xfs(children, xfs) do
    new_section = cell_xfs_element(xfs)

    {updated, found?} =
      Enum.map_reduce(children, false, fn
        {"cellXfs", _, _}, _ -> {new_section, true}
        other, seen -> {other, seen}
      end)

    if found?, do: updated, else: insert_cell_xfs(updated, new_section)
  end

  defp insert_cell_xfs(children, new_section) do
    idx =
      Enum.find_index(children, fn
        {"cellStyles", _, _} -> true
        {"dxfs", _, _} -> true
        {"tableStyles", _, _} -> true
        _ -> false
      end) || length(children)

    List.insert_at(children, idx, new_section)
  end

  defp cell_xfs_element(xfs) do
    children = Enum.map(xfs, &cell_format_element/1)
    {"cellXfs", [{"count", Integer.to_string(length(xfs))}], children}
  end

  defp cell_format_element(%CellFormat{} = xf) do
    attrs =
      [
        {"numFmtId", Integer.to_string(xf.num_fmt_id)},
        {"fontId", Integer.to_string(xf.font_id)},
        {"fillId", Integer.to_string(xf.fill_id)},
        {"borderId", Integer.to_string(xf.border_id)},
        {"xfId", Integer.to_string(xf.xf_id)}
      ] ++ apply_flags(xf)

    children =
      case xf.alignment do
        nil -> []
        alignment -> [alignment_element(alignment)]
      end

    {"xf", attrs, children}
  end

  defp apply_flags(%CellFormat{num_fmt_id: id, alignment: alignment}) do
    apply_number = if id != 0, do: [{"applyNumberFormat", "1"}], else: []
    apply_alignment = if alignment != nil, do: [{"applyAlignment", "1"}], else: []
    apply_number ++ apply_alignment
  end

  defp alignment_element(%AlignmentRecord{} = a) do
    attrs =
      []
      |> maybe_put("horizontal", to_string(a.horizontal), a.horizontal != :general)
      |> maybe_put("vertical", to_string(a.vertical), a.vertical != :bottom)
      |> maybe_put("wrapText", "1", a.wrap_text)
      |> maybe_put("textRotation", Integer.to_string(a.text_rotation), a.text_rotation != 0)
      |> maybe_put("indent", Integer.to_string(a.indent), a.indent != 0)
      |> maybe_put("shrinkToFit", "1", a.shrink_to_fit)

    {"alignment", attrs, []}
  end

  defp maybe_put(attrs, _name, _value, false), do: attrs
  defp maybe_put(attrs, name, value, true), do: attrs ++ [{name, value}]

  @doc """
  Adds (or finds) an `<xf>` with the given `numFmtId` sitting on top of the
  default font / fill / border. Returns `{index, updated_styles}` — the
  index is the value that goes into a cell's `s` attribute.
  """
  @spec upsert_date_format(t(), non_neg_integer()) :: {non_neg_integer(), t()}
  def upsert_date_format(%__MODULE__{cell_formats: xfs} = styles, num_fmt_id) do
    case Enum.find_index(xfs, fn xf ->
           xf.num_fmt_id == num_fmt_id and xf.font_id == 0 and xf.fill_id == 0 and
             xf.border_id == 0 and xf.alignment == nil
         end) do
      nil ->
        new_xf = %CellFormat{num_fmt_id: num_fmt_id}
        {length(xfs), %{styles | cell_formats: xfs ++ [new_xf]}}

      idx ->
        {idx, styles}
    end
  end

  @spec cell_format(t(), non_neg_integer()) :: {:ok, CellFormat.t()} | :error
  def cell_format(%__MODULE__{cell_formats: xfs}, index)
      when is_integer(index) and index >= 0 do
    case Enum.at(xfs, index) do
      nil -> :error
      xf -> {:ok, xf}
    end
  end

  @doc """
  Dereferences a cell's style index into a flat `%ExVEx.Style{}` record.
  Returns a default style for `nil` (an unstyled cell).
  """
  @spec resolve(t(), non_neg_integer() | nil) :: Style.t()
  def resolve(_styles, nil), do: %Style{}

  def resolve(%__MODULE__{} = styles, index) when is_integer(index) and index >= 0 do
    case cell_format(styles, index) do
      {:ok, xf} -> build_style(styles, xf)
      :error -> %Style{}
    end
  end

  @doc """
  Decides whether a cell format represents a date (or date-time / time).
  """
  @spec date_format?(t(), CellFormat.t()) :: boolean()
  def date_format?(%__MODULE__{num_fmts: num_fmts}, %CellFormat{num_fmt_id: id}) do
    cond do
      MapSet.member?(@builtin_date_ids, id) -> true
      id >= 164 -> custom_date_format?(num_fmts, id)
      true -> false
    end
  end

  @doc "The format code string for a given numFmt id, resolving built-ins."
  @spec format_code(t(), non_neg_integer()) :: String.t()
  def format_code(%__MODULE__{num_fmts: num_fmts}, id) when is_integer(id) do
    case Enum.find(num_fmts, &(&1.id == id)) do
      %NumFmt{format_code: code} -> code
      nil -> Map.get(@builtin_format_codes, id, "")
    end
  end

  defp custom_date_format?(num_fmts, id) do
    case Enum.find(num_fmts, &(&1.id == id)) do
      %NumFmt{format_code: code} -> date_like_code?(code)
      nil -> false
    end
  end

  defp date_like_code?(code) do
    stripped =
      code
      |> String.replace(~r/"[^"]*"/, "")
      |> String.replace(~r/\[[^\]]*\]/, "")

    Regex.match?(~r/[yYmMdDhHsS]/, stripped)
  end

  defp build_style(%__MODULE__{} = styles, %CellFormat{} = xf) do
    %Style{
      font: Enum.at(styles.fonts, xf.font_id, %Font{}),
      fill: Enum.at(styles.fills, xf.fill_id, %Fill{}),
      border: Enum.at(styles.borders, xf.border_id, %Border{}),
      alignment: to_alignment(xf.alignment),
      number_format: format_code(styles, xf.num_fmt_id)
    }
  end

  defp to_alignment(nil), do: %Alignment{}

  defp to_alignment(%AlignmentRecord{} = record) do
    %Alignment{
      horizontal: record.horizontal,
      vertical: record.vertical,
      wrap_text: record.wrap_text,
      text_rotation: record.text_rotation,
      indent: record.indent,
      shrink_to_fit: record.shrink_to_fit
    }
  end

  defp collect_num_fmts(children), do: section(children, "numFmts", &num_fmt_from_child/1)
  defp collect_fonts(children), do: section(children, "fonts", &font_from_child/1)
  defp collect_fills(children), do: section(children, "fills", &fill_from_child/1)
  defp collect_borders(children), do: section(children, "borders", &border_from_child/1)
  defp collect_cell_formats(children), do: section(children, "cellXfs", &cell_format_from_child/1)

  defp section(children, tag, builder) do
    children
    |> Enum.find_value([], fn
      {^tag, _, items} -> items
      _ -> nil
    end)
    |> Enum.flat_map(builder)
  end

  defp num_fmt_from_child({"numFmt", attrs, _}) do
    [
      %NumFmt{
        id: integer_attr!(attrs, "numFmtId"),
        format_code: fetch_attr!(attrs, "formatCode")
      }
    ]
  end

  defp num_fmt_from_child(_), do: []

  defp font_from_child({"font", _attrs, children}), do: [font_from_children(children)]
  defp font_from_child(_), do: []

  defp font_from_children(children) do
    Enum.reduce(children, %Font{}, fn
      {"sz", attrs, _}, font -> %{font | size: number_attr(attrs, "val", nil)}
      {"name", attrs, _}, font -> %{font | name: string_attr(attrs, "val", nil)}
      {"b", _, _}, font -> %{font | bold: true}
      {"i", _, _}, font -> %{font | italic: true}
      {"strike", _, _}, font -> %{font | strike: true}
      {"u", attrs, _}, font -> %{font | underline: underline_value(attrs)}
      {"color", attrs, _}, font -> %{font | color: color_from_attrs(attrs)}
      _, font -> font
    end)
  end

  @underline_values %{
    "single" => :single,
    "double" => :double,
    "singleAccounting" => :single_accounting,
    "doubleAccounting" => :double_accounting,
    "none" => :none
  }

  defp underline_value(attrs) do
    Map.get(@underline_values, string_attr(attrs, "val", "single"), :unknown)
  end

  defp color_from_attrs(attrs) do
    cond do
      rgb = List.keyfind(attrs, "rgb", 0) ->
        %Color{kind: :rgb, value: elem(rgb, 1)}

      theme = List.keyfind(attrs, "theme", 0) ->
        %Color{
          kind: :theme,
          value: integer_or_nil(elem(theme, 1)),
          tint: float_or_nil(attrs, "tint")
        }

      indexed = List.keyfind(attrs, "indexed", 0) ->
        %Color{kind: :indexed, value: integer_or_nil(elem(indexed, 1))}

      List.keyfind(attrs, "auto", 0) ->
        %Color{kind: :auto}

      true ->
        nil
    end
  end

  defp integer_or_nil(binary) do
    case Integer.parse(binary) do
      {n, ""} -> n
      _ -> nil
    end
  end

  defp float_or_nil(attrs, name) do
    case List.keyfind(attrs, name, 0) do
      {_, value} ->
        case Float.parse(value) do
          {f, _} -> f
          :error -> nil
        end

      nil ->
        nil
    end
  end

  defp fill_from_child({"fill", _attrs, children}) do
    case Enum.find(children, &match?({"patternFill", _, _}, &1)) do
      {"patternFill", attrs, pattern_children} ->
        [
          %Fill{
            pattern: pattern_type(attrs),
            foreground_color: color_child(pattern_children, "fgColor"),
            background_color: color_child(pattern_children, "bgColor")
          }
        ]

      nil ->
        [%Fill{}]
    end
  end

  defp fill_from_child(_), do: []

  @pattern_types %{
    "none" => :none,
    "solid" => :solid,
    "mediumGray" => :mediumGray,
    "darkGray" => :darkGray,
    "lightGray" => :lightGray,
    "darkHorizontal" => :darkHorizontal,
    "darkVertical" => :darkVertical,
    "darkDown" => :darkDown,
    "darkUp" => :darkUp,
    "darkGrid" => :darkGrid,
    "darkTrellis" => :darkTrellis,
    "lightHorizontal" => :lightHorizontal,
    "lightVertical" => :lightVertical,
    "lightDown" => :lightDown,
    "lightUp" => :lightUp,
    "lightGrid" => :lightGrid,
    "lightTrellis" => :lightTrellis,
    "gray125" => :gray125,
    "gray0625" => :gray0625
  }

  defp pattern_type(attrs) do
    case List.keyfind(attrs, "patternType", 0) do
      {_, value} -> Map.get(@pattern_types, value, :unknown)
      nil -> :none
    end
  end

  defp color_child(children, tag) do
    case Enum.find(children, &match?({^tag, _, _}, &1)) do
      {^tag, attrs, _} -> color_from_attrs(attrs)
      nil -> nil
    end
  end

  defp border_from_child({"border", _attrs, children}) do
    [
      %Border{
        top: side_child(children, "top"),
        bottom: side_child(children, "bottom"),
        left: side_child(children, "left"),
        right: side_child(children, "right")
      }
    ]
  end

  defp border_from_child(_), do: []

  defp side_child(children, tag) do
    case Enum.find(children, &match?({^tag, _, _}, &1)) do
      {^tag, attrs, inner} -> build_side(attrs, inner)
      nil -> %Side{}
    end
  end

  defp build_side(attrs, inner) do
    %Side{
      style: side_style(attrs),
      color: color_child(inner, "color")
    }
  end

  @side_styles %{
    "none" => :none,
    "thin" => :thin,
    "medium" => :medium,
    "dashed" => :dashed,
    "dotted" => :dotted,
    "thick" => :thick,
    "double" => :double,
    "hair" => :hair,
    "mediumDashed" => :mediumDashed,
    "dashDot" => :dashDot,
    "mediumDashDot" => :mediumDashDot,
    "dashDotDot" => :dashDotDot,
    "mediumDashDotDot" => :mediumDashDotDot,
    "slantDashDot" => :slantDashDot
  }

  defp side_style(attrs) do
    case List.keyfind(attrs, "style", 0) do
      {_, value} -> Map.get(@side_styles, value, :unknown)
      nil -> :none
    end
  end

  defp cell_format_from_child({"xf", attrs, children}) do
    [
      %CellFormat{
        num_fmt_id: integer_attr(attrs, "numFmtId", 0),
        font_id: integer_attr(attrs, "fontId", 0),
        fill_id: integer_attr(attrs, "fillId", 0),
        border_id: integer_attr(attrs, "borderId", 0),
        xf_id: integer_attr(attrs, "xfId", 0),
        alignment: alignment_from_children(children)
      }
    ]
  end

  defp cell_format_from_child(_), do: []

  defp alignment_from_children(children) do
    case Enum.find(children, &match?({"alignment", _, _}, &1)) do
      {"alignment", attrs, _} -> alignment_from_attrs(attrs)
      nil -> nil
    end
  end

  @horizontal_aligns %{
    "general" => :general,
    "left" => :left,
    "center" => :center,
    "right" => :right,
    "fill" => :fill,
    "justify" => :justify,
    "centerContinuous" => :center_continuous,
    "distributed" => :distributed
  }

  @vertical_aligns %{
    "top" => :top,
    "center" => :center,
    "bottom" => :bottom,
    "justify" => :justify,
    "distributed" => :distributed
  }

  defp alignment_from_attrs(attrs) do
    %AlignmentRecord{
      horizontal: enum_attr(attrs, "horizontal", @horizontal_aligns, :general),
      vertical: enum_attr(attrs, "vertical", @vertical_aligns, :bottom),
      wrap_text: bool_attr(attrs, "wrapText"),
      text_rotation: integer_attr(attrs, "textRotation", 0),
      indent: integer_attr(attrs, "indent", 0),
      shrink_to_fit: bool_attr(attrs, "shrinkToFit")
    }
  end

  defp enum_attr(attrs, name, mapping, default) do
    case List.keyfind(attrs, name, 0) do
      {_, value} -> Map.get(mapping, value, :unknown)
      nil -> default
    end
  end

  defp integer_attr(attrs, name, default) do
    case List.keyfind(attrs, name, 0) do
      {_, value} ->
        case Integer.parse(value) do
          {n, _} -> n
          _ -> default
        end

      nil ->
        default
    end
  end

  defp integer_attr!(attrs, name) do
    attrs |> fetch_attr!(name) |> String.to_integer()
  end

  defp number_attr(attrs, name, default) do
    case List.keyfind(attrs, name, 0) do
      {_, value} ->
        case Float.parse(value) do
          {f, _} -> f
          :error -> default
        end

      nil ->
        default
    end
  end

  defp string_attr(attrs, name, default) do
    case List.keyfind(attrs, name, 0) do
      {_, value} -> value
      nil -> default
    end
  end

  defp bool_attr(attrs, name) do
    case List.keyfind(attrs, name, 0) do
      {_, value} -> value in ["1", "true"]
      nil -> false
    end
  end

  defp fetch_attr!(attrs, name) do
    case List.keyfind(attrs, name, 0) do
      {^name, value} -> value
      nil -> raise ArgumentError, "missing #{name} attribute in styles.xml"
    end
  end
end
