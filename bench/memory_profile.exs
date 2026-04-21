defmodule Profile do
  @moduledoc false

  def word_size, do: :erlang.system_info(:wordsize)

  def term_bytes(term), do: :erts_debug.flat_size(term) * word_size()

  def bytes_to_mb(n), do: Float.round(n / 1_048_576, 2)

  def snapshot(label) do
    :erlang.garbage_collect()
    mem = :erlang.memory()

    IO.puts("\n=== #{label} ===")
    IO.puts("  Total BEAM:     #{bytes_to_mb(mem[:total])} MB")
    IO.puts("  Processes:      #{bytes_to_mb(mem[:processes])} MB")
    IO.puts("  ETS:            #{bytes_to_mb(mem[:ets])} MB")
    IO.puts("  Binary:         #{bytes_to_mb(mem[:binary])} MB")
    IO.puts("  Code:           #{bytes_to_mb(mem[:code])} MB")
    IO.puts("  Atom:           #{bytes_to_mb(mem[:atom])} MB")
    mem
  end

  def attribute_workbook(book, label) do
    IO.puts("\n--- Workbook internals: #{label} ---")

    parts_bytes =
      book.parts
      |> Map.values()
      |> Enum.reduce(0, fn bin, acc -> acc + byte_size(bin) end)

    IO.puts("  parts map (raw XML bytes):     #{bytes_to_mb(parts_bytes)} MB")

    largest_parts =
      book.parts
      |> Enum.map(fn {k, v} -> {k, byte_size(v)} end)
      |> Enum.sort_by(fn {_, s} -> -s end)
      |> Enum.take(5)

    IO.puts("  largest parts:")

    for {k, s} <- largest_parts do
      IO.puts("    #{String.pad_trailing(k, 40)}  #{bytes_to_mb(s)} MB")
    end

    sheet_ets_bytes =
      book.sheet_trees
      |> Map.values()
      |> Enum.reduce(0, fn editable, acc ->
        case editable.cells_table do
          nil -> acc
          t -> acc + :ets.info(t, :memory) * word_size()
        end
      end)

    IO.puts("  sheet cell ETS tables:         #{bytes_to_mb(sheet_ets_bytes)} MB")

    sst_ets_bytes =
      case book.shared_strings do
        nil ->
          0

        sst ->
          idx = if sst.by_index_table, do: :ets.info(sst.by_index_table, :memory), else: 0
          str = if sst.by_string_table, do: :ets.info(sst.by_string_table, :memory), else: 0
          (idx + str) * word_size()
      end

    IO.puts("  shared-string ETS tables:      #{bytes_to_mb(sst_ets_bytes)} MB")

    pre_post_bytes =
      book.sheet_trees
      |> Map.values()
      |> Enum.reduce(0, fn editable, acc ->
        acc + term_bytes(editable.pre_sheet_data) + term_bytes(editable.post_sheet_data)
      end)

    IO.puts("  sheet pre/post SimpleForm:     #{bytes_to_mb(pre_post_bytes)} MB")

    row_attrs_bytes =
      book.sheet_trees
      |> Map.values()
      |> Enum.reduce(0, fn editable, acc -> acc + term_bytes(editable.row_attrs) end)

    IO.puts("  row_attrs maps:                #{bytes_to_mb(row_attrs_bytes)} MB")

    structs =
      term_bytes(book.content_types) +
        term_bytes(book.workbook) +
        term_bytes(book.workbook_rels) +
        term_bytes(book.styles || %{})

    IO.puts("  parsed manifest/styles structs: #{bytes_to_mb(structs)} MB")

    total_accounted = parts_bytes + sheet_ets_bytes + sst_ets_bytes + pre_post_bytes + row_attrs_bytes + structs
    IO.puts("  --- attributed:                #{bytes_to_mb(total_accounted)} MB ---")
  end
end

spec_path = Path.join([File.cwd!(), "bench/comparative/spec.json"])
spec = spec_path |> File.read!() |> :json.decode()

# Use the same generator as bench/comparative/elixir/bench.exs
defmodule G do
  def gen_cell(row, 1), do: "SYM#{String.pad_leading(Integer.to_string(row), 4, "0")}"
  def gen_cell(row, 2), do: row * 1.5 + 2 * 0.25
  def gen_cell(row, 3), do: row * 0.1 + 3 * 0.01
  def gen_cell(row, 4), do: row * 4
  def gen_cell(row, 5), do: "LBL-#{String.pad_leading(Integer.to_string(row), 4, "0")}"
  def gen_cell(row, 6), do: rem(row, 2) == 0
  def gen_cell(_, 7), do: "B"
  def gen_cell(row, col), do: row + col
end

Profile.snapshot("Startup (before anything)")

summary_sheets = Map.fetch!(spec["template"], "summary_sheets")
data_sheets = Map.fetch!(spec["template"], "data_sheets")
reference_sheets = Map.fetch!(spec["template"], "reference_sheets")

{:ok, book} = ExVEx.new()
all_sheets = summary_sheets ++ data_sheets ++ reference_sheets

book =
  all_sheets
  |> Enum.with_index()
  |> Enum.reduce(book, fn {sheet, idx}, acc ->
    name = sheet["name"]

    if idx == 0 do
      {:ok, a} = ExVEx.rename_sheet(acc, "Sheet1", name)
      a
    else
      {:ok, a} = ExVEx.add_sheet(acc, name)
      a
    end
  end)

Profile.snapshot("After new + 12 sheets added")
Profile.attribute_workbook(book, "After new + 12 sheets added")

# Populate all data and reference sheets (the benchmark's create workload)
book =
  Enum.reduce(data_sheets ++ reference_sheets, book, fn sheet, acc ->
    name = sheet["name"]
    rows = sheet["rows"]
    cols = sheet["cols"]

    Enum.reduce(1..rows, acc, fn r, a ->
      Enum.reduce(1..cols, a, fn c, inner ->
        {:ok, next} = ExVEx.put_cell(inner, name, {r, c}, G.gen_cell(r, c))
        next
      end)
    end)
  end)

# Summary merges + formulas
book =
  Enum.reduce(summary_sheets, book, fn sheet, acc ->
    name = sheet["name"]

    acc =
      Enum.reduce(sheet["merges"] || [], acc, fn range, a ->
        {:ok, n} = ExVEx.merge_cells(a, name, range, preserve_values: true)
        n
      end)

    Enum.reduce(sheet["formulas"] || [], acc, fn [cell, formula], a ->
      {:ok, n} = ExVEx.put_cell(a, name, cell, {:formula, formula})
      n
    end)
  end)

Profile.snapshot("After populating all 24,800 cells")
Profile.attribute_workbook(book, "After populating all 24,800 cells")

out = Path.join(System.tmp_dir!(), "memory_profile_#{:erlang.unique_integer([:positive])}.xlsx")
:ok = ExVEx.save(book, out)

Profile.snapshot("After save")
Profile.attribute_workbook(book, "After save")

File.rm(out)
ExVEx.close(book)

Profile.snapshot("After close + rm")
