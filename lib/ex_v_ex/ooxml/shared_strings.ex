defmodule ExVEx.OOXML.SharedStrings do
  @moduledoc """
  The shared string table (`xl/sharedStrings.xml`).

  Most cell text in a real workbook is stored by index into this table — a
  cell with `t="s"` has a `<v>N</v>` that points at the Nth `<si>` entry.

  Storage is two maps: `by_index` and `by_string`. Both lookups and interns
  are O(1); interns are also allocation-cheap (map insertion, no copy of
  an N-element array as a tuple-backed table would need).

  Rich-text formatting (`<r>` runs, `<rPr>` properties) is flattened into
  plain text at parse time. Phonetic annotations (`<phoneticPr>`, `<rPh>`)
  are stripped from the parsed view. The original XML bytes remain in the
  workbook's `parts` map so untouched rich content round-trips even when
  ExVEx doesn't model it.
  """

  @namespace "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  @type t :: %__MODULE__{
          count: non_neg_integer(),
          by_index: %{non_neg_integer() => String.t()},
          by_string: %{String.t() => non_neg_integer()}
        }

  defstruct count: 0, by_index: %{}, by_string: %{}

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"sst", _attrs, children}} ->
        {:ok, build(Enum.flat_map(children, &string_item/1))}

      {:ok, _other} ->
        {:error, :not_a_shared_strings_table}

      {:error, reason} ->
        {:error, reason}
    end
  end

  @spec get(t(), integer()) :: {:ok, String.t()} | :error
  def get(%__MODULE__{by_index: by_index}, index)
      when is_integer(index) and index >= 0 do
    case Map.fetch(by_index, index) do
      {:ok, text} -> {:ok, text}
      :error -> :error
    end
  end

  def get(%__MODULE__{}, _), do: :error

  @spec count(t()) :: non_neg_integer()
  def count(%__MODULE__{count: n}), do: n

  @doc """
  Returns the index of a string in the table, adding it if it isn't there
  yet. Returns `{index, updated_sst}`. O(1).
  """
  @spec intern(t(), String.t()) :: {non_neg_integer(), t()}
  def intern(%__MODULE__{by_string: by_string} = sst, text) when is_binary(text) do
    case Map.fetch(by_string, text) do
      {:ok, existing} ->
        {existing, sst}

      :error ->
        new_index = sst.count

        {new_index,
         %{
           sst
           | count: new_index + 1,
             by_index: Map.put(sst.by_index, new_index, text),
             by_string: Map.put(by_string, text, new_index)
         }}
    end
  end

  @doc "Serialises the table back to an `xl/sharedStrings.xml` document."
  @spec serialize(t()) :: binary()
  def serialize(%__MODULE__{count: count, by_index: by_index}) do
    items =
      for i <- 0..(count - 1)//1 do
        text = Map.fetch!(by_index, i)
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
    {count, by_index, by_string} =
      Enum.reduce(strings, {0, %{}, %{}}, fn text, {i, idx_map, str_map} ->
        {i + 1, Map.put(idx_map, i, text), Map.put_new(str_map, text, i)}
      end)

    %__MODULE__{count: count, by_index: by_index, by_string: by_string}
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
