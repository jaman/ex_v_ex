defmodule ExVEx do
  @moduledoc """
  Pure-Elixir reader and editor for `.xlsx` / `.xlsm` workbooks.

  ## Quick start

      {:ok, book} = ExVEx.open("path/to/file.xlsx")
      ExVEx.sheet_names(book)            #=> ["Sheet1", "Sheet2"]
      {:ok, "hello"} = ExVEx.get_cell(book, "Sheet1", "A1")
      :ok = ExVEx.save(book, "path/to/output.xlsx")

  ## Design

  ExVEx opens a workbook lazily: raw ZIP part bytes are kept in memory and
  untouched parts are written back verbatim on `save/2`. This preserves
  unknown content (custom XML, VBA macros, extension schemas) on round-trip
  without the caller needing to opt in.
  """

  alias ExVEx.OOXML.SharedStrings
  alias ExVEx.OOXML.Styles
  alias ExVEx.OOXML.Workbook, as: WorkbookXml
  alias ExVEx.OOXML.Worksheet.Editable
  alias ExVEx.Packaging.{ContentTypes, Relationships, Zip}
  alias ExVEx.Utils.{Coordinate, Range}
  alias ExVEx.Workbook

  @type path :: Path.t()
  @type sheet_name :: String.t()
  @type cell_ref :: String.t() | {pos_integer(), pos_integer()}
  @type cell_value ::
          binary()
          | number()
          | boolean()
          | nil
          | {:formula, String.t()}
          | {:formula, String.t(), binary() | number() | boolean()}

  @package_rels_path "_rels/.rels"
  @office_document_type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  @shared_strings_type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  @styles_type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

  @spec open(path()) :: {:ok, Workbook.t()} | {:error, term()}
  def open(path) do
    with {:ok, entries} <- Zip.read(path) do
      parts = entries_to_parts(entries)
      part_order = Enum.map(entries, & &1.path)
      build_workbook(parts, part_order, path)
    end
  end

  @doc """
  Returns a minimal blank workbook with a single empty sheet named
  `"Sheet1"`. Compose with `add_sheet/2`, `rename_sheet/3`,
  `remove_sheet/2`, and `put_cell/4` to build a template from scratch.
  """
  @spec new() :: {:ok, Workbook.t()} | {:error, term()}
  def new do
    parts = %{
      "[Content_Types].xml" => blank_content_types(),
      "_rels/.rels" => blank_package_rels(),
      "xl/workbook.xml" => blank_workbook(),
      "xl/_rels/workbook.xml.rels" => blank_workbook_rels(),
      "xl/worksheets/sheet1.xml" => blank_sheet(),
      "xl/styles.xml" => blank_styles(),
      "xl/sharedStrings.xml" => blank_shared_strings()
    }

    part_order = [
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
      "xl/worksheets/sheet1.xml",
      "xl/styles.xml",
      "xl/sharedStrings.xml"
    ]

    build_workbook(parts, part_order, nil)
  end

  defp build_workbook(parts, part_order, source_path) do
    with {:ok, manifest_xml} <- fetch_part(parts, "[Content_Types].xml"),
         {:ok, content_types} <- ContentTypes.parse(manifest_xml),
         {:ok, package_rels_xml} <- fetch_part(parts, @package_rels_path),
         {:ok, package_rels} <- Relationships.parse(package_rels_xml),
         {:ok, workbook_path} <- resolve_workbook_path(package_rels),
         {:ok, workbook_xml} <- fetch_part(parts, workbook_path),
         {:ok, workbook} <- WorkbookXml.parse(workbook_xml),
         workbook_rels_path = rels_path_for(workbook_path),
         {:ok, workbook_rels_xml} <- fetch_part(parts, workbook_rels_path),
         {:ok, workbook_rels} <- Relationships.parse(workbook_rels_xml),
         sst_path = resolve_rel_target(workbook_rels, @shared_strings_type, workbook_path),
         styles_path = resolve_rel_target(workbook_rels, @styles_type, workbook_path),
         {:ok, shared_strings} <- maybe_load_part(parts, sst_path, &SharedStrings.parse/1),
         {:ok, styles} <- maybe_load_part(parts, styles_path, &Styles.parse/1) do
      {:ok,
       %Workbook{
         parts: parts,
         part_order: part_order,
         content_types: content_types,
         workbook: workbook,
         workbook_rels: workbook_rels,
         workbook_path: workbook_path,
         shared_strings: shared_strings,
         shared_strings_path: sst_path,
         styles: styles,
         styles_path: styles_path,
         source_path: source_path
       }}
    end
  end

  defp blank_content_types do
    ~s|<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>|
  end

  defp blank_package_rels do
    ~s|<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>|
  end

  defp blank_workbook do
    ~s|<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>|
  end

  defp blank_workbook_rels do
    ~s|<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>|
  end

  defp blank_shared_strings do
    ~s|<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>|
  end

  defp blank_sheet do
    ~s|<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>|
  end

  defp blank_styles do
    ~s|<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>|
  end

  @spec save(Workbook.t(), path()) :: :ok | {:error, term()}
  def save(%Workbook{} = book, path) do
    book
    |> Workbook.flush()
    |> Workbook.to_entries()
    |> then(&Zip.write(path, &1))
  end

  @doc """
  Releases the ETS tables backing a workbook's cell grid. Calling this is
  optional — tables are cleaned up when the owning process exits — but
  recommended in long-running processes that open many workbooks.

  After `close/1`, the workbook is no longer usable; calls to `get_cell`,
  `put_cell`, etc. will fail.
  """
  @spec close(Workbook.t()) :: :ok
  def close(%Workbook{} = book) do
    Enum.each(book.sheet_trees, fn {_path, %Editable{cells_table: t}} ->
      if t != nil, do: :ets.delete(t)
    end)

    if book.shared_strings, do: SharedStrings.close(book.shared_strings)

    :ok
  end

  @spec sheet_names(Workbook.t()) :: [sheet_name()]
  def sheet_names(%Workbook{workbook: %WorkbookXml{sheets: sheets}}) do
    Enum.map(sheets, & &1.name)
  end

  @worksheet_content_type "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  @worksheet_rel_type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

  @doc """
  Appends a new empty sheet with the given name to the workbook.

  Returns `{:error, :duplicate_sheet_name}` if the name is already in use.
  """
  @spec add_sheet(Workbook.t(), sheet_name()) :: {:ok, Workbook.t()} | {:error, term()}
  def add_sheet(%Workbook{} = book, name) when is_binary(name) do
    if sheet_exists?(book, name) do
      {:error, :duplicate_sheet_name}
    else
      {:ok, do_add_sheet(book, name)}
    end
  end

  defp do_add_sheet(book, name) do
    sheet_number = next_sheet_number(book)
    rel_id = next_rel_id(book.workbook_rels)
    sheet_path = "xl/worksheets/sheet#{sheet_number}.xml"

    new_sheet = %WorkbookXml.SheetRef{
      name: name,
      sheet_id: next_sheet_id(book),
      rel_id: rel_id,
      state: :visible
    }

    new_relationship = %Relationships.Relationship{
      id: rel_id,
      type: @worksheet_rel_type,
      target: "worksheets/sheet#{sheet_number}.xml"
    }

    new_override = %ContentTypes.Override{
      part_name: "/#{sheet_path}",
      content_type: @worksheet_content_type
    }

    book
    |> put_in(
      [Access.key(:workbook), Access.key(:sheets)],
      book.workbook.sheets ++ [new_sheet]
    )
    |> put_in(
      [Access.key(:workbook_rels), Access.key(:entries)],
      book.workbook_rels.entries ++ [new_relationship]
    )
    |> put_in(
      [Access.key(:content_types), Access.key(:overrides)],
      book.content_types.overrides ++ [new_override]
    )
    |> Map.update!(:parts, &Map.put(&1, sheet_path, blank_sheet()))
    |> Map.update!(:part_order, &(&1 ++ [sheet_path]))
    |> flush_package_metadata()
    |> Map.put(:calc_dirty, true)
  end

  @doc "Renames a sheet. Returns `{:error, :unknown_sheet}` if `old` is not found."
  @spec rename_sheet(Workbook.t(), sheet_name(), sheet_name()) ::
          {:ok, Workbook.t()} | {:error, term()}
  def rename_sheet(%Workbook{} = book, old, new) when is_binary(old) and is_binary(new) do
    cond do
      not sheet_exists?(book, old) ->
        {:error, :unknown_sheet}

      old == new ->
        {:ok, book}

      sheet_exists?(book, new) ->
        {:error, :duplicate_sheet_name}

      true ->
        new_sheets =
          Enum.map(book.workbook.sheets, fn
            %{name: ^old} = s -> %{s | name: new}
            s -> s
          end)

        book =
          book
          |> put_in([Access.key(:workbook), Access.key(:sheets)], new_sheets)
          |> flush_package_metadata()
          |> Map.put(:calc_dirty, true)

        {:ok, book}
    end
  end

  @doc """
  Removes a sheet and its worksheet part from the workbook.

  Returns `{:error, :unknown_sheet}` if the name is not found, or
  `{:error, :last_sheet}` if removing it would leave the workbook with
  zero sheets (invalid per the OOXML spec).
  """
  @spec remove_sheet(Workbook.t(), sheet_name()) :: {:ok, Workbook.t()} | {:error, term()}
  def remove_sheet(%Workbook{} = book, name) when is_binary(name) do
    cond do
      not sheet_exists?(book, name) ->
        {:error, :unknown_sheet}

      length(book.workbook.sheets) == 1 ->
        {:error, :last_sheet}

      true ->
        sheet_ref = Enum.find(book.workbook.sheets, &(&1.name == name))
        {:ok, path} = sheet_path(book, name)

        new_sheets = Enum.reject(book.workbook.sheets, &(&1.name == name))

        new_rels =
          Enum.reject(book.workbook_rels.entries, &(&1.id == sheet_ref.rel_id))

        new_overrides =
          Enum.reject(book.content_types.overrides, &(&1.part_name == "/" <> path))

        book =
          book
          |> put_in([Access.key(:workbook), Access.key(:sheets)], new_sheets)
          |> put_in([Access.key(:workbook_rels), Access.key(:entries)], new_rels)
          |> put_in([Access.key(:content_types), Access.key(:overrides)], new_overrides)
          |> Map.update!(:parts, &Map.delete(&1, path))
          |> Map.update!(:part_order, &Enum.reject(&1, fn p -> p == path end))
          |> Map.update!(:sheet_trees, &Map.delete(&1, path))
          |> Map.update!(:dirty_sheet_paths, &MapSet.delete(&1, path))
          |> flush_package_metadata()
          |> Map.put(:calc_dirty, true)

        {:ok, book}
    end
  end

  defp sheet_exists?(%Workbook{workbook: %{sheets: sheets}}, name) do
    Enum.any?(sheets, &(&1.name == name))
  end

  defp next_sheet_number(%Workbook{parts: parts}) do
    existing =
      parts
      |> Map.keys()
      |> Enum.flat_map(fn
        "xl/worksheets/sheet" <> rest ->
          case Integer.parse(rest) do
            {n, ".xml"} -> [n]
            _ -> []
          end

        _ ->
          []
      end)

    (Enum.max([0 | existing]) + 1) |> max(1)
  end

  defp next_sheet_id(%Workbook{workbook: %{sheets: sheets}}) do
    sheets |> Enum.map(& &1.sheet_id) |> Enum.max(fn -> 0 end) |> Kernel.+(1)
  end

  defp next_rel_id(%Relationships{entries: entries}) do
    max_n = entries |> Enum.flat_map(&parse_rel_id_number/1) |> Enum.max(fn -> 0 end)
    "rId#{max_n + 1}"
  end

  defp parse_rel_id_number(%Relationships.Relationship{id: "rId" <> rest}) do
    case Integer.parse(rest) do
      {n, ""} -> [n]
      _ -> []
    end
  end

  defp parse_rel_id_number(_), do: []

  defp flush_package_metadata(%Workbook{} = book) do
    workbook_xml =
      WorkbookXml.serialize_into(book.workbook, Map.fetch!(book.parts, book.workbook_path))

    rels_xml = Relationships.serialize(book.workbook_rels)
    rels_path = rels_path_for(book.workbook_path)

    ct_xml = ContentTypes.serialize(book.content_types)

    %{
      book
      | parts:
          book.parts
          |> Map.put(book.workbook_path, workbook_xml)
          |> Map.put(rels_path, rels_xml)
          |> Map.put("[Content_Types].xml", ct_xml)
    }
  end

  @spec sheet_path(Workbook.t(), sheet_name()) :: {:ok, String.t()} | :error
  def sheet_path(%Workbook{} = book, name) do
    with %{} = ref <- Enum.find(book.workbook.sheets, &(&1.name == name)),
         {:ok, rel} <- Relationships.get(book.workbook_rels, ref.rel_id) do
      {:ok, Relationships.resolve(rel, rels_path_for(book.workbook_path))}
    else
      _ -> :error
    end
  end

  @spec get_cell(Workbook.t(), sheet_name(), cell_ref()) ::
          {:ok, cell_value() | Date.t() | NaiveDateTime.t()} | {:error, term()}
  def get_cell(%Workbook{} = book, sheet, ref) do
    with {:ok, coord} <- parse_coordinate(ref),
         {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, _book} <- Workbook.fetch_sheet_tree(book, path) do
      case Editable.cell_record_at(editable, coord) do
        {:ok, cell} -> resolve_cell_value(cell, book)
        :error -> {:ok, nil}
      end
    end
  end

  @spec put_cell(Workbook.t(), sheet_name(), cell_ref(), cell_value()) ::
          {:ok, Workbook.t()} | {:error, term()}
  def put_cell(%Workbook{} = book, sheet, ref, value) do
    with {:ok, coord} <- parse_coordinate(ref),
         {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, book} <- Workbook.fetch_sheet_tree(book, path) do
      {encoded, book} = prepare_cell_value(book, value)
      new_editable = Editable.put_cell(editable, coord, encoded)
      {:ok, Workbook.put_sheet_tree(book, path, new_editable)}
    end
  end

  defp prepare_cell_value(%Workbook{shared_strings: %SharedStrings{}} = book, value)
       when is_binary(value) do
    {index, sst} = SharedStrings.intern(book.shared_strings, value)
    {{:shared_string, index}, %{book | shared_strings: sst, shared_strings_dirty: true}}
  end

  defp prepare_cell_value(book, %Date{} = date) do
    prepare_styled_serial(book, Date.to_gregorian_days(date) - gregorian_epoch(), 14)
  end

  defp prepare_cell_value(book, %NaiveDateTime{} = dt) do
    days = Date.to_gregorian_days(NaiveDateTime.to_date(dt)) - gregorian_epoch()
    {hours, minutes, seconds} = {dt.hour, dt.minute, dt.second}
    fraction = (hours * 3600 + minutes * 60 + seconds) / 86_400
    prepare_styled_serial(book, days + fraction, 22)
  end

  defp prepare_cell_value(book, value), do: {value, book}

  defp prepare_styled_serial(book, serial, num_fmt_id) do
    styles = book.styles || %Styles{}
    {style_id, styles} = Styles.upsert_date_format(styles, num_fmt_id)

    book = %{book | styles: styles, styles_dirty: true, styles_path: styles_path(book)}

    {{:styled, serial, style_id}, book}
  end

  defp styles_path(%Workbook{styles_path: path}) when is_binary(path), do: path
  defp styles_path(_), do: "xl/styles.xml"

  defp gregorian_epoch, do: Date.to_gregorian_days(~D[1899-12-30])

  @spec get_style(Workbook.t(), sheet_name(), cell_ref()) ::
          {:ok, ExVEx.Style.t()} | {:error, term()}
  def get_style(%Workbook{} = book, sheet, ref) do
    with {:ok, coord} <- parse_coordinate(ref),
         {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, _book} <- Workbook.fetch_sheet_tree(book, path) do
      style_id =
        case Editable.cell_record_at(editable, coord) do
          {:ok, %{style_id: id}} -> id
          :error -> nil
        end

      {:ok, Styles.resolve(book.styles || %Styles{}, style_id)}
    end
  end

  @type range_ref :: String.t()

  @doc """
  Merges a rectangular range of cells on a sheet.

  ## Options

    * `:preserve_values` — `false` (default) to clear every non-anchor cell
      in the range (Excel's convention); `true` to leave underlying cells
      untouched. Excel will still only display the anchor cell's value,
      but `get_cell/3` on a non-anchor cell will keep returning whatever
      was there.
    * `:on_overlap` — `:error` (default) to refuse a range that overlaps
      an existing merge and return `{:error, {:overlaps, existing_ref}}`;
      `:replace` to remove the overlapping range(s) first; `:allow` to
      permit overlapping ranges (matches openpyxl's lenient behaviour).
  """
  @spec merge_cells(Workbook.t(), sheet_name(), range_ref(), keyword()) ::
          {:ok, Workbook.t()} | {:error, term()}
  def merge_cells(%Workbook{} = book, sheet, ref, opts \\ []) do
    preserve = Keyword.get(opts, :preserve_values, false)
    on_overlap = Keyword.get(opts, :on_overlap, :error)

    with {:ok, range} <- parse_range(ref),
         {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, book} <- Workbook.fetch_sheet_tree(book, path),
         {:ok, editable} <- handle_overlap_on_editable(editable, range, on_overlap) do
      new_editable = Editable.merge(editable, range, not preserve)
      {:ok, Workbook.put_sheet_tree(book, path, new_editable)}
    end
  end

  @doc """
  Removes a merged range from a sheet.

  ## Options

    * `:on_missing` — `:error` (default) returns `{:error, :not_merged}`
      when the exact range is not currently merged; `:ignore` makes the
      call a no-op in that case.
  """
  @spec unmerge_cells(Workbook.t(), sheet_name(), range_ref(), keyword()) ::
          {:ok, Workbook.t()} | {:error, term()}
  def unmerge_cells(%Workbook{} = book, sheet, ref, opts \\ []) do
    on_missing = Keyword.get(opts, :on_missing, :error)

    with {:ok, range} <- parse_range(ref),
         {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, book} <- Workbook.fetch_sheet_tree(book, path) do
      existing = Editable.merged_ranges(editable)
      apply_unmerge(book, path, editable, range, existing, on_missing)
    end
  end

  defp apply_unmerge(book, path, editable, range, existing, on_missing) do
    if Enum.any?(existing, &exact_match?(&1, range)) do
      new_editable = Editable.unmerge(editable, range)
      {:ok, Workbook.put_sheet_tree(book, path, new_editable)}
    else
      handle_missing_unmerge(book, on_missing)
    end
  end

  @doc """
  Returns the list of merged ranges on a sheet as A1-style range refs.
  """
  @spec merged_ranges(Workbook.t(), sheet_name()) :: {:ok, [range_ref()]} | {:error, term()}
  def merged_ranges(%Workbook{} = book, sheet) do
    with {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, _book} <- Workbook.fetch_sheet_tree(book, path) do
      {:ok, editable |> Editable.merged_ranges() |> Enum.map(&Range.to_string/1)}
    end
  end

  defp parse_range(ref) do
    case Range.parse(ref) do
      {:ok, range} -> {:ok, range}
      :error -> {:error, :invalid_range}
    end
  end

  defp exact_match?(%Range{} = a, %Range{} = b) do
    a.top_left == b.top_left and a.bottom_right == b.bottom_right
  end

  defp handle_overlap_on_editable(editable, _range, :allow), do: {:ok, editable}

  defp handle_overlap_on_editable(editable, range, mode) when mode in [:error, :replace] do
    existing = Editable.merged_ranges(editable)
    overlap = Enum.find(existing, &(Range.overlaps?(&1, range) and not exact_match?(&1, range)))

    case {overlap, mode} do
      {nil, _} -> {:ok, editable}
      {conflict, :error} -> {:error, {:overlaps, Range.to_string(conflict)}}
      {conflict, :replace} -> {:ok, Editable.unmerge(editable, conflict)}
    end
  end

  defp handle_missing_unmerge(book, :ignore), do: {:ok, book}
  defp handle_missing_unmerge(_book, :error), do: {:error, :not_merged}

  @spec get_formula(Workbook.t(), sheet_name(), cell_ref()) ::
          {:ok, String.t() | nil} | {:error, term()}
  def get_formula(%Workbook{} = book, sheet, ref) do
    with {:ok, coord} <- parse_coordinate(ref),
         {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, _book} <- Workbook.fetch_sheet_tree(book, path) do
      case Editable.cell_record_at(editable, coord) do
        {:ok, %{formula: formula}} -> {:ok, formula}
        :error -> {:ok, nil}
      end
    end
  end

  @doc """
  Returns every populated cell on a sheet as a map of A1 refs to resolved
  values. Empty cells are omitted.
  """
  @spec cells(Workbook.t(), sheet_name()) :: {:ok, %{cell_ref() => term()}} | {:error, term()}
  def cells(%Workbook{} = book, sheet) do
    with {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, _book} <- Workbook.fetch_sheet_tree(book, path) do
      cells = Editable.cells_map(editable)
      {:ok, Enum.reduce(cells, %{}, &put_resolved(&1, &2, book))}
    end
  end

  defp put_resolved({coord, cell}, acc, book) do
    case resolve_cell_value(cell, book) do
      {:ok, value} -> Map.put(acc, Coordinate.to_string(coord), value)
      {:error, _} -> acc
    end
  end

  @doc """
  Streams every populated cell on a sheet as `{a1_ref, value}` pairs in
  row-major order (row 1 before row 2; column A before column B within a
  row).
  """
  @spec each_cell(Workbook.t(), sheet_name()) ::
          {:ok, Enumerable.t({cell_ref(), term()})} | {:error, term()}
  def each_cell(%Workbook{} = book, sheet) do
    with {:ok, path} <- sheet_path_or_error(book, sheet),
         {:ok, editable, _book} <- Workbook.fetch_sheet_tree(book, path) do
      cells = Editable.cells_map(editable)

      stream =
        cells
        |> Enum.sort_by(fn {coord, _} -> coord end)
        |> Stream.flat_map(&resolve_to_pair(&1, book))

      {:ok, stream}
    end
  end

  defp resolve_to_pair({coord, cell}, book) do
    case resolve_cell_value(cell, book) do
      {:ok, value} -> [{Coordinate.to_string(coord), value}]
      {:error, _} -> []
    end
  end

  defp sheet_path_or_error(book, sheet) do
    case sheet_path(book, sheet) do
      {:ok, path} -> {:ok, path}
      :error -> {:error, :unknown_sheet}
    end
  end

  defp parse_coordinate({row, col})
       when is_integer(row) and row > 0 and is_integer(col) and col > 0 do
    {:ok, {row, col}}
  end

  defp parse_coordinate(ref) when is_binary(ref) do
    case Coordinate.parse(ref) do
      {:ok, coord} -> {:ok, coord}
      :error -> {:error, :invalid_coordinate}
    end
  end

  defp parse_coordinate(_), do: {:error, :invalid_coordinate}

  defp resolve_cell_value(%{raw_type: :shared_string, raw_value: idx_str}, %Workbook{
         shared_strings: %SharedStrings{} = sst
       }) do
    with {idx, ""} <- Integer.parse(idx_str || ""),
         {:ok, text} <- SharedStrings.get(sst, idx) do
      {:ok, text}
    else
      _ -> {:error, :invalid_shared_string_index}
    end
  end

  defp resolve_cell_value(%{raw_type: :shared_string}, _), do: {:error, :no_shared_string_table}

  defp resolve_cell_value(%{raw_type: :inline_string, raw_value: text}, _), do: {:ok, text || ""}
  defp resolve_cell_value(%{raw_type: :boolean, raw_value: "1"}, _), do: {:ok, true}
  defp resolve_cell_value(%{raw_type: :boolean, raw_value: "0"}, _), do: {:ok, false}

  defp resolve_cell_value(%{raw_type: :number} = cell, %Workbook{} = book) do
    case parse_number(cell.raw_value) do
      {:ok, number} when is_number(number) -> maybe_as_date(number, cell, book)
      other -> other
    end
  end

  defp resolve_cell_value(%{raw_type: :formula_string, raw_value: text}, _), do: {:ok, text || ""}

  defp resolve_cell_value(%{raw_type: :error, raw_value: code}, _) do
    {:error, {:cell_error, code}}
  end

  defp maybe_as_date(number, %{style_id: nil}, _), do: {:ok, number}
  defp maybe_as_date(number, _cell, %Workbook{styles: nil}), do: {:ok, number}

  defp maybe_as_date(number, %{style_id: style_id}, %Workbook{styles: styles}) do
    with {:ok, xf} <- Styles.cell_format(styles, style_id),
         true <- Styles.date_format?(styles, xf),
         {:ok, value} <- serial_to_temporal(number, date_format_code(styles, xf)) do
      {:ok, value}
    else
      _ -> {:ok, number}
    end
  end

  defp date_format_code(%Styles{} = styles, %{num_fmt_id: id}), do: Styles.format_code(styles, id)

  defp serial_to_temporal(number, format_code) do
    if time_bearing_code?(format_code) do
      serial_to_naive_datetime(number)
    else
      serial_to_date(number)
    end
  end

  defp time_bearing_code?(code), do: Regex.match?(~r/[hHsS]/, code)

  defp serial_to_date(serial) when is_number(serial) and serial >= 1 do
    days = trunc(serial)
    epoch = if days < 60, do: ~D[1899-12-31], else: ~D[1899-12-30]
    {:ok, Date.add(epoch, days)}
  end

  defp serial_to_date(_), do: {:error, :serial_out_of_range}

  defp serial_to_naive_datetime(serial) when is_number(serial) and serial >= 0 do
    days = trunc(serial)
    fraction = serial - days
    seconds_in_day = round(fraction * 86_400)

    date =
      if days == 0 do
        ~D[1899-12-30]
      else
        epoch = if days < 60, do: ~D[1899-12-31], else: ~D[1899-12-30]
        Date.add(epoch, days)
      end

    {:ok, NaiveDateTime.add(NaiveDateTime.new!(date, ~T[00:00:00]), seconds_in_day, :second)}
  end

  defp serial_to_naive_datetime(_), do: {:error, :serial_out_of_range}

  defp parse_number(nil), do: {:ok, nil}

  defp parse_number(raw) do
    if String.contains?(raw, [".", "e", "E"]) do
      case Float.parse(raw) do
        {f, ""} -> {:ok, f}
        _ -> {:error, {:bad_number, raw}}
      end
    else
      case Integer.parse(raw) do
        {n, ""} -> {:ok, n}
        _ -> {:error, {:bad_number, raw}}
      end
    end
  end

  defp entries_to_parts(entries) do
    Map.new(entries, fn %{path: p, data: data} -> {p, data} end)
  end

  defp fetch_part(parts, key) do
    case Map.fetch(parts, key) do
      {:ok, data} -> {:ok, data}
      :error -> {:error, {:missing_part, key}}
    end
  end

  defp resolve_workbook_path(%Relationships{entries: entries}) do
    case Enum.find(entries, &(&1.type == @office_document_type)) do
      nil -> {:error, :no_office_document_relationship}
      rel -> {:ok, Relationships.resolve(rel, @package_rels_path)}
    end
  end

  defp resolve_rel_target(%Relationships{entries: entries}, type, workbook_path) do
    case Enum.find(entries, &(&1.type == type)) do
      nil -> nil
      rel -> Relationships.resolve(rel, rels_path_for(workbook_path))
    end
  end

  defp maybe_load_part(_parts, nil, _parser), do: {:ok, nil}

  defp maybe_load_part(parts, path, parser) do
    case Map.fetch(parts, path) do
      {:ok, xml} -> parser.(xml)
      :error -> {:ok, nil}
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
end
