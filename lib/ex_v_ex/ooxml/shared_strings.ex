defmodule ExVEx.OOXML.SharedStrings do
  @moduledoc """
  The shared string table (`xl/sharedStrings.xml`).

  Most cell text in a real workbook is stored by index into this table — a
  cell with `t="s"` has a `<v>N</v>` that points at the Nth `<si>` entry.

  ExVEx stores the table as a tuple so lookup is O(1). Rich-text formatting
  (`<r>` runs, `<rPr>` properties) is flattened into plain text at the
  moment of parse; the original XML bytes are preserved in the workbook's
  `parts` map so rich text round-trips even if we don't understand it.

  Phonetic annotations (`<phoneticPr>`, `<rPh>`) are stripped from the
  parsed view for the same reason.
  """

  @type t :: %__MODULE__{strings: tuple()}
  defstruct strings: {}

  @spec parse(binary()) :: {:ok, t()} | {:error, term()}
  def parse(xml) when is_binary(xml) do
    case Saxy.SimpleForm.parse_string(xml) do
      {:ok, {"sst", _attrs, children}} ->
        strings = children |> Enum.flat_map(&string_item/1) |> List.to_tuple()
        {:ok, %__MODULE__{strings: strings}}

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
