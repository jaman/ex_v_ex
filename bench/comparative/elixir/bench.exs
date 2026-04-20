[spec_path, mode, input, output] = System.argv()

spec =
  spec_path
  |> File.read!()
  |> :json.decode()

defmodule Bench do
  def generate_data_row(row, cols, _seed) do
    for col <- 1..cols do
      {col, gen_cell(row, col)}
    end
  end

  defp gen_cell(row, 1), do: "SYM#{String.pad_leading(Integer.to_string(row), 4, "0")}"
  defp gen_cell(row, 2), do: row * 1.5 + 2 * 0.25
  defp gen_cell(row, 3), do: row * 0.1 + 3 * 0.01
  defp gen_cell(row, 4), do: row * 4
  defp gen_cell(row, 5), do: "LBL-#{String.pad_leading(Integer.to_string(row), 4, "0")}"
  defp gen_cell(row, 6), do: rem(row, 2) == 0
  defp gen_cell(_row, 7), do: "B"
  defp gen_cell(row, col), do: row + col

  def write_data_sheet(book, sheet_name, rows, cols) do
    Enum.reduce(1..rows, book, fn row, acc ->
      Enum.reduce(1..cols, acc, fn col, inner ->
        {:ok, next} = ExVEx.put_cell(inner, sheet_name, {row, col}, gen_cell(row, col))
        next
      end)
    end)
  end

  def clear_data_sheet(book, sheet_name, rows, cols) do
    Enum.reduce(1..rows, book, fn row, acc ->
      Enum.reduce(1..cols, acc, fn col, inner ->
        case ExVEx.put_cell(inner, sheet_name, {row, col}, nil) do
          {:ok, next} -> next
          {:error, _} -> inner
        end
      end)
    end)
  end
end

summary_sheets = Map.fetch!(spec["template"], "summary_sheets")
data_sheets = Map.fetch!(spec["template"], "data_sheets")
reference_sheets = Map.fetch!(spec["template"], "reference_sheets")

{wall_us, {cells_written, cells_cleared}} =
  :timer.tc(fn ->
    case mode do
      "create" ->
        {:ok, book} = ExVEx.new()

        all_sheets = summary_sheets ++ data_sheets ++ reference_sheets

        book =
          all_sheets
          |> Enum.with_index()
          |> Enum.reduce(book, fn {sheet, idx}, acc ->
            name = sheet["name"]

            acc =
              if idx == 0 do
                {:ok, a} = ExVEx.rename_sheet(acc, "Sheet1", name)
                a
              else
                {:ok, a} = ExVEx.add_sheet(acc, name)
                a
              end

            acc
          end)

        # populate data sheets
        {book, written} =
          Enum.reduce(data_sheets, {book, 0}, fn sheet, {acc, count} ->
            name = sheet["name"]
            rows = sheet["rows"]
            cols = sheet["cols"]
            {Bench.write_data_sheet(acc, name, rows, cols), count + rows * cols}
          end)

        # populate reference sheets (smaller)
        {book, written} =
          Enum.reduce(reference_sheets, {book, written}, fn sheet, {acc, count} ->
            name = sheet["name"]
            rows = sheet["rows"]
            cols = sheet["cols"]
            {Bench.write_data_sheet(acc, name, rows, cols), count + rows * cols}
          end)

        # apply summary sheet merges + formulas
        book =
          Enum.reduce(summary_sheets, book, fn sheet, acc ->
            name = sheet["name"]

            acc =
              Enum.reduce(sheet["merges"] || [], acc, fn range, a ->
                {:ok, next} = ExVEx.merge_cells(a, name, range, preserve_values: true)
                next
              end)

            Enum.reduce(sheet["formulas"] || [], acc, fn [cell, formula], a ->
              {:ok, next} = ExVEx.put_cell(a, name, cell, {:formula, formula})
              next
            end)
          end)

        :ok = ExVEx.save(book, output)
        {written, 0}

      "edit" ->
        {:ok, book} = ExVEx.open(input)

        data_sheet_names =
          data_sheets
          |> Enum.map(& &1["name"])
          |> Enum.filter(&Enum.member?(ExVEx.sheet_names(book), &1))

        rows = spec["manipulation"]["rows_per_data_sheet"]
        cols = spec["manipulation"]["cols_per_data_sheet"]

        {book, cleared} =
          Enum.reduce(data_sheet_names, {book, 0}, fn name, {acc, n} ->
            {Bench.clear_data_sheet(acc, name, rows, cols), n + rows * cols}
          end)

        {book, written} =
          Enum.reduce(data_sheet_names, {book, 0}, fn name, {acc, n} ->
            {Bench.write_data_sheet(acc, name, rows, cols), n + rows * cols}
          end)

        :ok = ExVEx.save(book, output)
        {written, cleared}
    end
  end)

result = %{
  "lang" => "elixir",
  "mode" => mode,
  "wall_ms" => Float.round(wall_us / 1000, 3),
  "cells_written" => cells_written,
  "cells_cleared" => cells_cleared
}

IO.puts(:json.encode(result))
