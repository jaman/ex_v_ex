defmodule ExVEx.OOXML.Styles do
  @moduledoc """
  Parser for `xl/styles.xml`.

  The OOXML style model is an indirection graph: a cell carries a style
  index `s="N"` into `cellXfs`; each `<xf>` there points at a `numFmt`,
  `font`, `fill`, and `border` by index. ExVEx currently models only the
  pieces it needs to reason about cell values — the number format and the
  font reference — and leaves fills/borders/alignment as opaque indices
  that round-trip through the raw `parts` map.

  Built-in number format IDs (0..163) are not listed in `numFmts` at all;
  workbook-specific custom formats start at ID 164.
  """

  alias ExVEx.OOXML.Styles.{CellFormat, NumFmt}

  @type t :: %__MODULE__{
          num_fmts: [NumFmt.t()],
          cell_formats: [CellFormat.t()]
        }

  defstruct num_fmts: [], cell_formats: []

  @builtin_date_ids MapSet.new([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47])

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"styleSheet", _attrs, children}} ->
        {:ok,
         %__MODULE__{
           num_fmts: collect_num_fmts(children),
           cell_formats: collect_cell_formats(children)
         }}

      {:ok, _other} ->
        {:error, :not_a_stylesheet}

      {:error, reason} ->
        {:error, reason}
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
  Decides whether a cell format represents a date (or date-time / time).

  Returns `true` for any `%CellFormat{}` whose `num_fmt_id` is either a
  built-in date format (14-22, 45-47) or a custom format whose code
  contains date tokens like `y`, `m`, `d`, `h`, `s`.
  """
  @spec date_format?(t(), CellFormat.t()) :: boolean()
  def date_format?(%__MODULE__{num_fmts: num_fmts}, %CellFormat{num_fmt_id: id}) do
    cond do
      MapSet.member?(@builtin_date_ids, id) ->
        true

      id >= 164 ->
        case Enum.find(num_fmts, &(&1.id == id)) do
          %NumFmt{format_code: code} -> date_like_code?(code)
          nil -> false
        end

      true ->
        false
    end
  end

  defp date_like_code?(code) do
    stripped = strip_literal_sections(code)
    Regex.match?(~r/[yYmMdDhHsS]/, stripped)
  end

  defp strip_literal_sections(code) do
    code
    |> String.replace(~r/"[^"]*"/, "")
    |> String.replace(~r/\[[^\]]*\]/, "")
  end

  defp collect_num_fmts(children) do
    children
    |> Enum.find_value([], fn
      {"numFmts", _, items} -> items
      _ -> nil
    end)
    |> Enum.flat_map(&num_fmt_from_child/1)
  end

  defp num_fmt_from_child({"numFmt", attrs, _}) do
    [
      %NumFmt{
        id: attrs |> fetch_attr!("numFmtId") |> String.to_integer(),
        format_code: fetch_attr!(attrs, "formatCode")
      }
    ]
  end

  defp num_fmt_from_child(_), do: []

  defp collect_cell_formats(children) do
    children
    |> Enum.find_value([], fn
      {"cellXfs", _, items} -> items
      _ -> nil
    end)
    |> Enum.flat_map(&cell_format_from_child/1)
  end

  defp cell_format_from_child({"xf", attrs, _}) do
    [
      %CellFormat{
        num_fmt_id: integer_attr(attrs, "numFmtId", 0),
        font_id: integer_attr(attrs, "fontId", 0),
        fill_id: integer_attr(attrs, "fillId", 0),
        border_id: integer_attr(attrs, "borderId", 0),
        xf_id: integer_attr(attrs, "xfId", 0)
      }
    ]
  end

  defp cell_format_from_child(_), do: []

  defp integer_attr(attrs, name, default) do
    case List.keyfind(attrs, name, 0) do
      {_, value} -> String.to_integer(value)
      nil -> default
    end
  end

  defp fetch_attr!(attrs, name) do
    case List.keyfind(attrs, name, 0) do
      {^name, value} -> value
      nil -> raise ArgumentError, "missing #{name} attribute in styles.xml"
    end
  end
end
