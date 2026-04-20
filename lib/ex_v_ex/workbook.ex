defmodule ExVEx.Workbook do
  @moduledoc """
  The in-memory representation of a workbook.

  A workbook holds every part from the source archive as raw bytes (the
  `parts` map) plus a parsed view of the handful of parts ExVEx actually
  reasons about (the `content_types` manifest, and eventually the
  relationships graph, sheet list, shared strings, and stylesheet).

  Parts that higher-level APIs have not yet parsed — worksheets not yet
  accessed, custom XML, VBA binaries, extension data — pass through
  untouched on `save/2`. This is how ExVEx preserves round-trip fidelity
  without exhaustively modelling every OOXML schema.
  """

  alias ExVEx.OOXML.{SharedStrings, Styles, Worksheet}
  alias ExVEx.OOXML.Workbook, as: WorkbookXml
  alias ExVEx.Packaging.ContentTypes
  alias ExVEx.Packaging.Relationships
  alias ExVEx.Packaging.Zip.Entry

  @type parts :: %{String.t() => binary()}

  @enforce_keys [:parts, :part_order, :content_types, :workbook, :workbook_rels, :workbook_path]
  defstruct parts: %{},
            part_order: [],
            content_types: %ContentTypes{},
            workbook: %WorkbookXml{},
            workbook_rels: %Relationships{},
            workbook_path: "xl/workbook.xml",
            shared_strings: nil,
            shared_strings_path: nil,
            shared_strings_dirty: false,
            styles: nil,
            styles_path: nil,
            styles_dirty: false,
            calc_dirty: false,
            sheet_trees: %{},
            dirty_sheet_paths: MapSet.new(),
            source_path: nil

  @type t :: %__MODULE__{
          parts: %{String.t() => binary()},
          part_order: [String.t()],
          content_types: ContentTypes.t(),
          workbook: WorkbookXml.t(),
          workbook_rels: Relationships.t(),
          workbook_path: String.t(),
          shared_strings: SharedStrings.t() | nil,
          shared_strings_path: String.t() | nil,
          shared_strings_dirty: boolean(),
          styles: Styles.t() | nil,
          styles_path: String.t() | nil,
          styles_dirty: boolean(),
          calc_dirty: boolean(),
          sheet_trees: %{String.t() => tuple()},
          dirty_sheet_paths: MapSet.t(String.t()),
          source_path: Path.t() | nil
        }

  @doc """
  Returns the parsed worksheet tree for `path`, parsing and caching it on
  first access. Subsequent calls reuse the cached tree so bulk reads and
  writes don't pay the parse cost repeatedly.
  """
  @spec fetch_sheet_tree(t(), String.t()) :: {:ok, tuple(), t()} | {:error, term()}
  def fetch_sheet_tree(%__MODULE__{sheet_trees: trees} = book, path) do
    case Map.fetch(trees, path) do
      {:ok, tree} ->
        {:ok, tree, book}

      :error ->
        with {:ok, xml} <- fetch_part(book.parts, path),
             {:ok, tree} <- Worksheet.parse_tree(xml) do
          {:ok, tree, %{book | sheet_trees: Map.put(trees, path, tree)}}
        end
    end
  end

  @doc """
  Updates the cached tree for `path` and marks the sheet dirty so
  `flush/1` re-serializes it on save.
  """
  @spec put_sheet_tree(t(), String.t(), tuple()) :: t()
  def put_sheet_tree(%__MODULE__{} = book, path, tree) do
    %{
      book
      | sheet_trees: Map.put(book.sheet_trees, path, tree),
        dirty_sheet_paths: MapSet.put(book.dirty_sheet_paths, path),
        calc_dirty: true
    }
  end

  defp fetch_part(parts, key) do
    case Map.fetch(parts, key) do
      {:ok, data} -> {:ok, data}
      :error -> {:error, {:missing_part, key}}
    end
  end

  @doc """
  Produces a new workbook whose `parts` map has every in-memory parsed
  sub-document (shared strings, styles) re-serialized and written back into
  the raw parts map. Called by `ExVEx.save/2` right before ZIP emission.
  """
  @spec flush(t()) :: t()
  def flush(%__MODULE__{} = book) do
    book
    |> flush_sheet_trees()
    |> flush_shared_strings()
    |> flush_styles()
    |> flush_calc_invalidation()
  end

  defp flush_sheet_trees(%__MODULE__{dirty_sheet_paths: paths} = book) do
    if MapSet.size(paths) == 0 do
      book
    else
      Enum.reduce(paths, book, fn path, acc ->
        tree = Map.fetch!(acc.sheet_trees, path)
        xml = Worksheet.encode_tree(tree)
        %{acc | parts: Map.put(acc.parts, path, xml)}
      end)
      |> Map.put(:dirty_sheet_paths, MapSet.new())
    end
  end

  defp flush_calc_invalidation(%__MODULE__{calc_dirty: false} = book), do: book

  defp flush_calc_invalidation(%__MODULE__{} = book) do
    book
    |> drop_calc_chain()
    |> force_full_calc_on_load()
    |> Map.put(:calc_dirty, false)
  end

  defp drop_calc_chain(%__MODULE__{parts: parts} = book) do
    calc_chain_path = find_calc_chain_path(book)

    if calc_chain_path do
      book
      |> Map.put(:parts, Map.delete(parts, calc_chain_path))
      |> Map.put(:part_order, Enum.reject(book.part_order, &(&1 == calc_chain_path)))
      |> drop_calc_chain_content_type()
      |> drop_calc_chain_relationship(calc_chain_path)
    else
      book
    end
  end

  defp find_calc_chain_path(%__MODULE__{parts: parts}) do
    Enum.find(Map.keys(parts), &String.ends_with?(&1, "calcChain.xml"))
  end

  defp drop_calc_chain_content_type(%__MODULE__{content_types: ct, parts: parts} = book) do
    new_overrides =
      Enum.reject(ct.overrides, &String.ends_with?(&1.part_name, "calcChain.xml"))

    new_ct = %{ct | overrides: new_overrides}

    %{
      book
      | content_types: new_ct,
        parts: Map.put(parts, "[Content_Types].xml", ContentTypes.serialize(new_ct))
    }
  end

  defp drop_calc_chain_relationship(%__MODULE__{workbook_rels: rels, parts: parts} = book, _path) do
    new_entries =
      Enum.reject(rels.entries, &String.ends_with?(&1.target, "calcChain.xml"))

    new_rels = %{rels | entries: new_entries}
    rels_path = rels_path_for(book.workbook_path)

    %{
      book
      | workbook_rels: new_rels,
        parts: Map.put(parts, rels_path, Relationships.serialize(new_rels))
    }
  end

  defp force_full_calc_on_load(%__MODULE__{parts: parts, workbook_path: path} = book) do
    case Map.fetch(parts, path) do
      {:ok, xml} ->
        case WorkbookXml.set_full_calc_on_load(xml) do
          {:ok, new_xml} -> %{book | parts: Map.put(parts, path, new_xml)}
          {:error, _} -> book
        end

      :error ->
        book
    end
  end

  defp rels_path_for(part_path) do
    dir = Path.dirname(part_path)
    base = Path.basename(part_path)

    case dir do
      "." -> "_rels/#{base}.rels"
      _ -> "#{dir}/_rels/#{base}.rels"
    end
  end

  defp flush_shared_strings(%__MODULE__{shared_strings_dirty: false} = book), do: book

  defp flush_shared_strings(%__MODULE__{shared_strings: sst, shared_strings_path: path} = book)
       when not is_nil(sst) and not is_nil(path) do
    xml = SharedStrings.serialize(sst)
    %{book | parts: Map.put(book.parts, path, xml), shared_strings_dirty: false}
  end

  defp flush_shared_strings(book), do: %{book | shared_strings_dirty: false}

  defp flush_styles(%__MODULE__{styles_dirty: false} = book), do: book

  defp flush_styles(%__MODULE__{styles: styles, styles_path: path} = book)
       when not is_nil(styles) and not is_nil(path) do
    original = Map.get(book.parts, path, "")
    xml = Styles.serialize_into(styles, original)
    %{book | parts: Map.put(book.parts, path, xml), styles_dirty: false}
  end

  defp flush_styles(book), do: %{book | styles_dirty: false}

  @doc """
  Projects a workbook back into the flat list of ZIP entries needed by
  `ExVEx.Packaging.Zip.write/2`.

  Parts are written in original source order; any parts added after `open/1`
  land at the end. Every part is emitted from its raw bytes in `parts`, so
  untouched content is byte-preserved on round-trip.
  """
  @spec to_entries(t()) :: [Entry.t()]
  def to_entries(%__MODULE__{parts: parts, part_order: order}) do
    seen = MapSet.new(order)
    ordered = for path <- order, Map.has_key?(parts, path), do: entry(parts, path)
    added = for {path, _} <- parts, not MapSet.member?(seen, path), do: entry(parts, path)
    ordered ++ added
  end

  defp entry(parts, path) do
    %Entry{path: path, data: Map.fetch!(parts, path)}
  end
end
