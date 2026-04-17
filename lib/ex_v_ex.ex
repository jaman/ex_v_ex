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

  alias ExVEx.OOXML.{SharedStrings, Styles, Worksheet}
  alias ExVEx.OOXML.Workbook, as: WorkbookXml
  alias ExVEx.Packaging.{ContentTypes, Relationships, Zip}
  alias ExVEx.Utils.{Coordinate, Range}
  alias ExVEx.Workbook

  @type path :: Path.t()
  @type sheet_name :: String.t()
  @type cell_ref :: String.t()
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
    with {:ok, entries} <- Zip.read(path),
         parts = entries_to_parts(entries),
         {:ok, manifest_xml} <- fetch_part(parts, "[Content_Types].xml"),
         {:ok, content_types} <- ContentTypes.parse(manifest_xml),
         {:ok, package_rels_xml} <- fetch_part(parts, @package_rels_path),
         {:ok, package_rels} <- Relationships.parse(package_rels_xml),
         {:ok, workbook_path} <- resolve_workbook_path(package_rels),
         {:ok, workbook_xml} <- fetch_part(parts, workbook_path),
         {:ok, workbook} <- WorkbookXml.parse(workbook_xml),
         workbook_rels_path = rels_path_for(workbook_path),
         {:ok, workbook_rels_xml} <- fetch_part(parts, workbook_rels_path),
         {:ok, workbook_rels} <- Relationships.parse(workbook_rels_xml),
         {:ok, shared_strings, sst_path} <-
           maybe_load_shared_strings(parts, workbook_rels, workbook_path),
         {:ok, styles, styles_path} <- maybe_load_styles(parts, workbook_rels, workbook_path) do
      {:ok,
       %Workbook{
         parts: parts,
         part_order: Enum.map(entries, & &1.path),
         content_types: content_types,
         workbook: workbook,
         workbook_rels: workbook_rels,
         workbook_path: workbook_path,
         shared_strings: shared_strings,
         shared_strings_path: sst_path,
         styles: styles,
         styles_path: styles_path,
         source_path: path
       }}
    end
  end

  @spec save(Workbook.t(), path()) :: :ok | {:error, term()}
  def save(%Workbook{} = book, path) do
    book
    |> Workbook.flush()
    |> Workbook.to_entries()
    |> then(&Zip.write(path, &1))
  end

  @spec sheet_names(Workbook.t()) :: [sheet_name()]
  def sheet_names(%Workbook{workbook: %WorkbookXml{sheets: sheets}}) do
    Enum.map(sheets, & &1.name)
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
         {:ok, sheet_xml} <- fetch_sheet_xml(book, sheet),
         {:ok, worksheet} <- Worksheet.parse(sheet_xml) do
      case Map.fetch(worksheet.cells, coord) do
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
         {:ok, xml} <- fetch_part(book.parts, path) do
      {encoded, book} = prepare_cell_value(book, value)

      with {:ok, new_xml} <- Worksheet.put_cell(xml, coord, encoded) do
        {:ok, %{book | parts: Map.put(book.parts, path, new_xml)}}
      end
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
         {:ok, sheet_xml} <- fetch_sheet_xml(book, sheet),
         {:ok, worksheet} <- Worksheet.parse(sheet_xml) do
      style_id =
        case Map.fetch(worksheet.cells, coord) do
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
         {:ok, xml} <- fetch_part(book.parts, path),
         {:ok, xml} <- handle_overlap(xml, range, on_overlap),
         {:ok, new_xml} <- Worksheet.merge(xml, range, not preserve) do
      {:ok, %{book | parts: Map.put(book.parts, path, new_xml)}}
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
         {:ok, xml} <- fetch_part(book.parts, path),
         {:ok, existing} <- Worksheet.merged_ranges(xml) do
      apply_unmerge(book, path, xml, range, existing, on_missing)
    end
  end

  defp apply_unmerge(book, path, xml, range, existing, on_missing) do
    if Enum.any?(existing, &exact_match?(&1, range)) do
      with {:ok, new_xml} <- Worksheet.unmerge(xml, range),
           do: {:ok, %{book | parts: Map.put(book.parts, path, new_xml)}}
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
         {:ok, xml} <- fetch_part(book.parts, path),
         {:ok, ranges} <- Worksheet.merged_ranges(xml) do
      {:ok, Enum.map(ranges, &Range.to_string/1)}
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

  defp handle_overlap(xml, _range, :allow), do: {:ok, xml}

  defp handle_overlap(xml, range, mode) when mode in [:error, :replace] do
    {:ok, existing} = Worksheet.merged_ranges(xml)
    overlap = Enum.find(existing, &(Range.overlaps?(&1, range) and not exact_match?(&1, range)))

    case {overlap, mode} do
      {nil, _} ->
        {:ok, xml}

      {conflict, :error} ->
        {:error, {:overlaps, Range.to_string(conflict)}}

      {conflict, :replace} ->
        Worksheet.unmerge(xml, conflict)
    end
  end

  defp handle_missing_unmerge(book, :ignore), do: {:ok, book}
  defp handle_missing_unmerge(_book, :error), do: {:error, :not_merged}

  @spec get_formula(Workbook.t(), sheet_name(), cell_ref()) ::
          {:ok, String.t() | nil} | {:error, term()}
  def get_formula(%Workbook{} = book, sheet, ref) do
    with {:ok, coord} <- parse_coordinate(ref),
         {:ok, sheet_xml} <- fetch_sheet_xml(book, sheet),
         {:ok, worksheet} <- Worksheet.parse(sheet_xml) do
      case Map.fetch(worksheet.cells, coord) do
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
    with {:ok, sheet_xml} <- fetch_sheet_xml(book, sheet),
         {:ok, worksheet} <- Worksheet.parse(sheet_xml) do
      {:ok, Enum.reduce(worksheet.cells, %{}, &put_resolved(&1, &2, book))}
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
    with {:ok, sheet_xml} <- fetch_sheet_xml(book, sheet),
         {:ok, worksheet} <- Worksheet.parse(sheet_xml) do
      stream =
        worksheet.cells
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

  defp parse_coordinate(ref) do
    case Coordinate.parse(ref) do
      {:ok, coord} -> {:ok, coord}
      :error -> {:error, :invalid_coordinate}
    end
  end

  defp fetch_sheet_xml(book, sheet) do
    with {:ok, path} <- sheet_path_or_error(book, sheet) do
      fetch_part(book.parts, path)
    end
  end

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

  defp maybe_load_shared_strings(parts, rels, workbook_path) do
    load_relationship_part(
      parts,
      rels,
      workbook_path,
      @shared_strings_type,
      &SharedStrings.parse/1
    )
  end

  defp maybe_load_styles(parts, rels, workbook_path) do
    load_relationship_part(parts, rels, workbook_path, @styles_type, &Styles.parse/1)
  end

  defp load_relationship_part(
         parts,
         %Relationships{entries: entries},
         workbook_path,
         type,
         parser
       ) do
    case Enum.find(entries, &(&1.type == type)) do
      nil -> {:ok, nil, nil}
      rel -> load_found_part(parts, rel, workbook_path, parser)
    end
  end

  defp load_found_part(parts, rel, workbook_path, parser) do
    path = Relationships.resolve(rel, rels_path_for(workbook_path))

    case Map.fetch(parts, path) do
      {:ok, xml} -> with {:ok, parsed} <- parser.(xml), do: {:ok, parsed, path}
      :error -> {:ok, nil, nil}
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
