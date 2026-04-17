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

  alias ExVEx.OOXML.{SharedStrings, Styles}
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
            styles: nil,
            source_path: nil

  @type t :: %__MODULE__{
          parts: %{String.t() => binary()},
          part_order: [String.t()],
          content_types: ContentTypes.t(),
          workbook: WorkbookXml.t(),
          workbook_rels: Relationships.t(),
          workbook_path: String.t(),
          shared_strings: SharedStrings.t() | nil,
          styles: Styles.t() | nil,
          source_path: Path.t() | nil
        }

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
