label = System.get_env("BENCH_LABEL") || "baseline"
tag = System.get_env("BENCH_TAG") || label
load = System.get_env("BENCH_LOAD")

results_dir = Path.join([File.cwd!(), "bench", "results"])
File.mkdir_p!(results_dir)
save_path = Path.join(results_dir, "#{label}.benchee")

fixture = Path.join([File.cwd!(), "test", "fixtures", "cells.xlsx"])

put_cells_then_save = fn ->
  {:ok, book} = ExVEx.open(fixture)

  book =
    Enum.reduce(1..500, book, fn i, acc ->
      row = div(i - 1, 5) + 1
      col = rem(i - 1, 5) + 1
      {:ok, next} = ExVEx.put_cell(acc, "Sheet1", {row, col}, "value-#{i}")
      next
    end)

  out = Path.join(System.tmp_dir!(), "bench_put_cell_#{:erlang.unique_integer([:positive])}.xlsx")
  :ok = ExVEx.save(book, out)
  File.rm(out)
end

read_500_cells = fn ->
  {:ok, book} = ExVEx.open(fixture)

  Enum.each(1..500, fn i ->
    row = div(i - 1, 5) + 1
    col = rem(i - 1, 5) + 1
    {:ok, _} = ExVEx.get_cell(book, "Sheet1", {row, col})
  end)
end

intern_1000_unique_strings = fn ->
  {:ok, book} = ExVEx.open(fixture)

  book =
    Enum.reduce(1..1000, book, fn i, acc ->
      {:ok, next} = ExVEx.put_cell(acc, "Sheet1", {i, 1}, "s-#{i}")
      next
    end)

  out = Path.join(System.tmp_dir!(), "bench_sst_#{:erlang.unique_integer([:positive])}.xlsx")
  :ok = ExVEx.save(book, out)
  File.rm(out)
end

IO.puts("\n=== ExVEx benchmark — #{label} ===\n")

run_opts = [
  time: 5,
  warmup: 2,
  memory_time: 1,
  save: [path: save_path, tag: tag],
  formatters: [
    {Benchee.Formatters.Console, comparison: true, extended_statistics: false}
  ]
]

run_opts = if load, do: Keyword.put(run_opts, :load, load), else: run_opts

Benchee.run(
  %{
    "500 put_cell + save" => put_cells_then_save,
    "500 get_cell" => read_500_cells,
    "1000 unique string interns + save" => intern_1000_unique_strings
  },
  run_opts
)

IO.puts("\nSnapshot saved to #{save_path}")
