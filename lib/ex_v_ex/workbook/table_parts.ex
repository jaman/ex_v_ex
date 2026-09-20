defmodule ExVEx.Workbook.TableParts do
  @moduledoc """
  Locates, adds, rewrites, and removes the `xl/tables/table*.xml` parts
  of a workbook, keeping the worksheet's `<tableParts>` list, the
  worksheet `.rels` file, and `[Content_Types].xml` in step.

  Every function takes and returns a `%ExVEx.Workbook{}`. Table parts
  are parsed on each call; only parts this module writes are
  re-serialized.
  """

  alias ExVEx.Mutation.Shift, as: MutShift
  alias ExVEx.OOXML.Table
  alias ExVEx.OOXML.Worksheet.Editable
  alias ExVEx.Packaging.ContentTypes
  alias ExVEx.Packaging.ContentTypes.Override
  alias ExVEx.Packaging.Relationships
  alias ExVEx.Packaging.Relationships.Relationship
  alias ExVEx.Workbook

  @table_rel "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
  @table_content_type "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
  @r_ns "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  defmodule Entry do
    @moduledoc "A table part together with the sheet and relationship that own it."

    @enforce_keys [:sheet, :sheet_path, :rels_path, :rel_id, :part_path, :table]
    defstruct [:sheet, :sheet_path, :rels_path, :rel_id, :part_path, :table]

    @type t :: %__MODULE__{
            sheet: String.t(),
            sheet_path: String.t(),
            rels_path: String.t(),
            rel_id: String.t(),
            part_path: String.t(),
            table: Table.t()
          }
  end

  @spec list(Workbook.t()) :: [Entry.t()]
  def list(%Workbook{} = book) do
    Enum.flat_map(book.workbook.sheets, fn sheet_ref ->
      case Workbook.sheet_path(book, sheet_ref.name) do
        {:ok, sheet_path} -> sheet_entries(book, sheet_ref.name, sheet_path)
        :error -> []
      end
    end)
  end

  @spec sheet_entries(Workbook.t(), String.t(), String.t()) :: [Entry.t()]
  def sheet_entries(%Workbook{} = book, sheet_name, sheet_path) do
    rels_path = Relationships.rels_path_for(sheet_path)

    for rel <- rels(book, rels_path).entries,
        rel.type == @table_rel,
        rel.target_mode == :internal,
        part_path = Relationships.resolve(rel, rels_path),
        {:ok, xml} <- [Map.fetch(book.parts, part_path)],
        {:ok, table} <- [Table.parse(xml)] do
      %Entry{
        sheet: sheet_name,
        sheet_path: sheet_path,
        rels_path: rels_path,
        rel_id: rel.id,
        part_path: part_path,
        table: table
      }
    end
  end

  @spec fetch(Workbook.t(), String.t()) :: {:ok, Entry.t()} | :error
  def fetch(%Workbook{} = book, name) when is_binary(name) do
    wanted = String.downcase(name)

    case Enum.find(list(book), &(String.downcase(&1.table.name) == wanted)) do
      nil -> :error
      entry -> {:ok, entry}
    end
  end

  @spec next_id(Workbook.t()) :: pos_integer()
  def next_id(%Workbook{} = book) do
    book |> list() |> Enum.map(& &1.table.id) |> Enum.max(fn -> 0 end) |> Kernel.+(1)
  end

  @doc """
  Creates the part for `table` on the sheet at `sheet_path`, links it
  from the worksheet, and registers its content type. Returns the
  workbook and the new entry.
  """
  @spec add(Workbook.t(), String.t(), String.t(), Table.t()) ::
          {:ok, Entry.t(), Workbook.t()} | {:error, term()}
  def add(%Workbook{} = book, sheet_name, sheet_path, %Table{} = table) do
    with {:ok, editable, book} <- Workbook.fetch_sheet_tree(book, sheet_path) do
      part_number = next_part_number(book)
      part_path = "xl/tables/table#{part_number}.xml"
      rels_path = Relationships.rels_path_for(sheet_path)
      rels = rels(book, rels_path)
      rel_id = Relationships.next_id(rels)

      relationship = %Relationship{
        id: rel_id,
        type: @table_rel,
        target: "../tables/table#{part_number}.xml"
      }

      new_editable =
        editable
        |> Editable.ensure_namespace("xmlns:r", @r_ns)
        |> Editable.add_table_part(rel_id)

      book =
        book
        |> put_part(part_path, Table.serialize(table))
        |> put_part(rels_path, Relationships.serialize(Relationships.append(rels, relationship)))
        |> add_override(part_path)
        |> Workbook.put_sheet_tree(sheet_path, new_editable)

      entry = %Entry{
        sheet: sheet_name,
        sheet_path: sheet_path,
        rels_path: rels_path,
        rel_id: rel_id,
        part_path: part_path,
        table: table
      }

      {:ok, entry, book}
    end
  end

  @doc "Rewrites the part for `entry` with `table`."
  @spec put(Workbook.t(), Entry.t(), Table.t()) :: Workbook.t()
  def put(%Workbook{} = book, %Entry{part_path: part_path}, %Table{} = table) do
    book |> put_part(part_path, Table.serialize(table)) |> Map.put(:calc_dirty, true)
  end

  @doc "Deletes the part, its relationship, its `<tablePart>` entry, and its content type."
  @spec remove(Workbook.t(), Entry.t()) :: Workbook.t()
  def remove(%Workbook{} = book, %Entry{} = entry) do
    rels = Relationships.delete(rels(book, entry.rels_path), entry.rel_id)

    book =
      case rels.entries do
        [] -> delete_part(book, entry.rels_path)
        _ -> put_part(book, entry.rels_path, Relationships.serialize(rels))
      end

    book =
      case Workbook.fetch_sheet_tree(book, entry.sheet_path) do
        {:ok, editable, book} ->
          Workbook.put_sheet_tree(
            book,
            entry.sheet_path,
            Editable.remove_table_part(editable, entry.rel_id)
          )

        {:error, _} ->
          book
      end

    book
    |> delete_part(entry.part_path)
    |> remove_override(entry.part_path)
    |> Map.put(:calc_dirty, true)
  end

  @doc """
  Applies a structural shift on `sheet_path` to every table on that
  sheet. Tables the shift empties are removed; the rest are rewritten.
  Returns the workbook and, for each surviving table, the entry before
  and after the shift so the caller can reconcile cells.
  """
  @spec shift(Workbook.t(), String.t(), String.t(), MutShift.t()) ::
          {Workbook.t(), [{Entry.t(), Entry.t()}]}
  def shift(%Workbook{} = book, sheet_name, sheet_path, %MutShift{} = mut_shift) do
    book
    |> sheet_entries(sheet_name, sheet_path)
    |> Enum.reduce({book, []}, fn entry, {acc, survivors} ->
      case Table.shift(entry.table, mut_shift) do
        :deleted ->
          {remove(acc, entry), survivors}

        {:ok, shifted} ->
          new_entry = %{entry | table: shifted}
          {put(acc, entry, shifted), [{entry, new_entry} | survivors]}
      end
    end)
    |> then(fn {acc, survivors} -> {acc, Enum.reverse(survivors)} end)
  end

  defp rels(%Workbook{parts: parts}, rels_path) do
    with {:ok, xml} <- Map.fetch(parts, rels_path),
         {:ok, rels} <- Relationships.parse(xml) do
      rels
    else
      _ -> %Relationships{}
    end
  end

  defp next_part_number(%Workbook{parts: parts}) do
    parts
    |> Map.keys()
    |> Enum.flat_map(fn
      "xl/tables/table" <> rest ->
        case Integer.parse(rest) do
          {n, ".xml"} -> [n]
          _ -> []
        end

      _ ->
        []
    end)
    |> Enum.max(fn -> 0 end)
    |> Kernel.+(1)
  end

  defp put_part(%Workbook{parts: parts, part_order: order} = book, path, data) do
    new_order = if path in order, do: order, else: order ++ [path]
    %{book | parts: Map.put(parts, path, data), part_order: new_order}
  end

  defp delete_part(%Workbook{parts: parts, part_order: order} = book, path) do
    %{book | parts: Map.delete(parts, path), part_order: List.delete(order, path)}
  end

  defp add_override(%Workbook{content_types: ct} = book, part_path) do
    override = %Override{part_name: "/" <> part_path, content_type: @table_content_type}
    write_content_types(book, %{ct | overrides: ct.overrides ++ [override]})
  end

  defp remove_override(%Workbook{content_types: ct} = book, part_path) do
    part_name = "/" <> part_path

    write_content_types(book, %{
      ct
      | overrides: Enum.reject(ct.overrides, &(&1.part_name == part_name))
    })
  end

  defp write_content_types(%Workbook{} = book, %ContentTypes{} = ct) do
    %{book | content_types: ct}
    |> put_part("[Content_Types].xml", ContentTypes.serialize(ct))
  end
end
