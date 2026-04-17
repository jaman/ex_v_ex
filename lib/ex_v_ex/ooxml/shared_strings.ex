defmodule ExVEx.OOXML.SharedStrings do
  @moduledoc """
  The shared string table (`xl/sharedStrings.xml`).

  Most cell text in a real workbook is stored by index into this table — a
  cell with `t="s"` has a `<v>N</v>` that points at the Nth `<si>` entry.

  Lookups (`get/2`) are O(1) via a tuple. `intern/2` is O(log n) via a map
  index and returns the existing index if the string is already in the
  table, otherwise appends and returns the new index.

  Rich-text formatting (`<r>` runs, `<rPr>` properties) is flattened into
  plain text at parse time. Phonetic annotations (`<phoneticPr>`, `<rPh>`)
  are stripped from the parsed view. The original XML bytes remain in the
  workbook's `parts` map so untouched rich content round-trips even when
  ExVEx doesn't model it.
  """

  @namespace "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  @type t :: %__MODULE__{
          strings: tuple(),
          index: %{String.t() => non_neg_integer()}
        }

  defstruct strings: {}, index: %{}

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"sst", _attrs, children}} ->
        strings = children |> Enum.flat_map(&string_item/1)
        {:ok, build(strings)}

      {:ok, _other} ->
        {:error, :not_a_shared_strings_table}

      {:error, reason} ->
        {:error, reason}
    end
  end

  @spec get(t(), integer()) :: {:ok, String.t()} | :error
  def get(%__MODULE__{strings: strings}, index)
      when is_integer(index) and index >= 0 and index < tuple_size(strings) do
    {:ok, elem(strings, index)}
  end

  def get(%__MODULE__{}, _), do: :error

  @spec count(t()) :: non_neg_integer()
  def count(%__MODULE__{strings: strings}), do: tuple_size(strings)

  @doc """
  Returns the index of a string in the table, adding it if it isn't there
  yet. Returns `{index, updated_sst}`.
  """
  @spec intern(t(), String.t()) :: {non_neg_integer(), t()}
  def intern(%__MODULE__{index: index, strings: strings} = sst, text) when is_binary(text) do
    case Map.fetch(index, text) do
      {:ok, existing} ->
        {existing, sst}

      :error ->
        new_index = tuple_size(strings)
        new_strings = Tuple.insert_at(strings, tuple_size(strings), text)
        new_map = Map.put(index, text, new_index)
        {new_index, %{sst | strings: new_strings, index: new_map}}
    end
  end

  @doc "Serialises the table back to an `xl/sharedStrings.xml` document."
  @spec serialize(t()) :: binary()
  def serialize(%__MODULE__{strings: strings}) do
    count = tuple_size(strings)

    items =
      for i <- 0..(count - 1)//1 do
        text = elem(strings, i)
        Saxy.XML.element("si", [], [Saxy.XML.element("t", text_attrs(text), [text])])
      end

    root =
      Saxy.XML.element(
        "sst",
        [
          {"xmlns", @namespace},
          {"count", Integer.to_string(count)},
          {"uniqueCount", Integer.to_string(count)}
        ],
        items
      )

    Saxy.encode!(root, version: "1.0", encoding: "UTF-8", standalone: true)
  end

  defp text_attrs(text) do
    if String.match?(text, ~r/\A\s|\s\z|\n/), do: [{"xml:space", "preserve"}], else: []
  end

  defp build(strings) do
    index =
      strings
      |> Enum.with_index()
      |> Map.new()

    %__MODULE__{strings: List.to_tuple(strings), index: index}
  end

  defp string_item({"si", _attrs, children}) do
    [Enum.map_join(children, "", &text_from_si_child/1)]
  end

  defp string_item(_), do: []

  defp text_from_si_child({"t", _attrs, children}), do: text_content(children)
  defp text_from_si_child({"r", _attrs, children}), do: text_from_run(children)
  defp text_from_si_child(_), do: ""

  defp text_from_run(children) do
    Enum.map_join(children, "", fn
      {"t", _attrs, text_children} -> text_content(text_children)
      _ -> ""
    end)
  end

  defp text_content(children) do
    Enum.map_join(children, "", fn
      text when is_binary(text) -> text
      _ -> ""
    end)
  end
end
