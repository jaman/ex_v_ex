defmodule ExVEx.OOXML.SharedStrings do
  @moduledoc """
  The shared string table (`xl/sharedStrings.xml`).

  Most cell text in a real workbook is stored by index into this table — a
  cell with `t="s"` has a `<v>N</v>` that points at the Nth `<si>` entry.

  Storage is ETS-backed: one `:set` table for index→string, another for
  string→index. Lookups are O(1). Interns are O(1) and allocate nothing
  beyond the single ETS row.

  Because the tables are ETS, multiple `%ExVEx.Workbook{}` references
  derived from the same `ExVEx.open/1` or `ExVEx.new/0` share the same
  underlying SST state. This matches the sharing semantics of the
  per-sheet cell tables so that string-valued cell mutations remain
  internally consistent across workbook references.

  Rich-text formatting (`<r>` runs, `<rPr>` properties) is flattened into
  plain text at parse time. Phonetic annotations (`<phoneticPr>`,
  `<rPh>`) are stripped from the parsed view. The original XML bytes
  remain in the workbook's `parts` map so untouched rich content
  round-trips even when ExVEx doesn't model it.
  """

  @namespace "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  @type t :: %__MODULE__{
          by_index_table: :ets.tid() | nil,
          by_string_table: :ets.tid() | nil
        }

  defstruct by_index_table: nil, by_string_table: nil

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
  def get(%__MODULE__{by_index_table: table}, index)
      when is_integer(index) and index >= 0 and not is_nil(table) do
    case :ets.lookup(table, index) do
      [{^index, text}] -> {:ok, text}
      [] -> :error
    end
  end

  def get(%__MODULE__{}, _), do: :error

  @spec count(t()) :: non_neg_integer()
  def count(%__MODULE__{by_index_table: nil}), do: 0
  def count(%__MODULE__{by_index_table: table}), do: :ets.info(table, :size)

  @doc """
  Returns the index of a string in the table, adding it if it isn't there
  yet. Returns `{index, sst}` — `sst` is the same struct (ETS mutation is
  in-place), included so callers can treat this identically to the
  Map-backed version.
  """
  @spec intern(t(), String.t()) :: {non_neg_integer(), t()}
  def intern(%__MODULE__{by_string_table: str_table} = sst, text)
      when is_binary(text) and not is_nil(str_table) do
    case :ets.lookup(str_table, text) do
      [{^text, existing}] ->
        {existing, sst}

      [] ->
        new_index = :ets.info(sst.by_index_table, :size)
        :ets.insert(sst.by_index_table, {new_index, text})
        :ets.insert(str_table, {text, new_index})
        {new_index, sst}
    end
  end

  @doc "Serialises the table back to an `xl/sharedStrings.xml` document."
  @spec serialize(t()) :: binary()
  def serialize(%__MODULE__{by_index_table: nil}) do
    empty_root = Saxy.XML.element("sst", empty_sst_attrs(), [])
    Saxy.encode!(empty_root, version: "1.0", encoding: "UTF-8", standalone: true)
  end

  def serialize(%__MODULE__{by_index_table: table}) do
    count = :ets.info(table, :size)

    items =
      for i <- 0..(count - 1)//1 do
        [{^i, text}] = :ets.lookup(table, i)
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

  @doc """
  Releases the backing ETS tables. Call from `ExVEx.close/1` when done.
  """
  @spec close(t()) :: :ok
  def close(%__MODULE__{by_index_table: idx, by_string_table: str}) do
    if idx != nil, do: :ets.delete(idx)
    if str != nil, do: :ets.delete(str)
    :ok
  end

  defp empty_sst_attrs do
    [{"xmlns", @namespace}, {"count", "0"}, {"uniqueCount", "0"}]
  end

  defp text_attrs(text) do
    if String.match?(text, ~r/\A\s|\s\z|\n/), do: [{"xml:space", "preserve"}], else: []
  end

  defp build(strings) do
    by_index = :ets.new(:exvex_sst_by_index, [:set, :public])
    by_string = :ets.new(:exvex_sst_by_string, [:set, :public])

    _final =
      Enum.reduce(strings, 0, fn text, i ->
        :ets.insert(by_index, {i, text})
        :ets.insert_new(by_string, {text, i})
        i + 1
      end)

    %__MODULE__{by_index_table: by_index, by_string_table: by_string}
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
