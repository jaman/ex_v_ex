defmodule Mix.Tasks.Bench.Comparative do
  @moduledoc """
  Runs the rotating 4-phase xlsx benchmark across Elixir / Python / Rust.

  Each participating runtime must be on `PATH`:

  - `elixir` (always available — this is a mix task)
  - `python3` + openpyxl installed (pip install -r bench/comparative/python/requirements.txt)
  - `cargo` with a local Rust toolchain

  Missing runtimes are skipped with a warning. The report reflects only the
  runtimes that actually ran.

  Options:

    --skip-python   Skip the Python implementation even if available.
    --skip-rust     Skip the Rust implementation even if available.
  """
  use Mix.Task

  alias ExVEx.Bench.Report

  @shortdoc "Run the 4-phase cross-language xlsx benchmark"

  @impl Mix.Task
  def run(argv) do
    {opts, _, _} =
      OptionParser.parse(argv, switches: [skip_python: :boolean, skip_rust: :boolean])

    Mix.Task.run("app.config")

    base = Path.expand("bench/comparative", File.cwd!())
    spec_path = Path.join(base, "spec.json")

    langs = detect_langs(opts, base)

    if langs == [] do
      Mix.raise("No runtimes available. Need at least Elixir.")
    end

    run_dir = prepare_run_dir(base)

    IO.puts("=== Comparative benchmark — run #{Path.basename(run_dir)} ===")
    IO.puts("Runtimes: #{Enum.map_join(langs, ", ", & &1.name)}")
    IO.puts("Output:   #{run_dir}")
    IO.puts("")

    results = run_phases(langs, spec_path, run_dir)

    write_results(results, run_dir)
    generate_report(results, run_dir, langs)

    IO.puts("")
    IO.puts("Report: #{Path.join(run_dir, "report.md")}")
    IO.puts("Chart:  #{Path.join(run_dir, "report.html")}")
  end

  defp detect_langs(opts, base) do
    [
      elixir_lang(),
      python_lang(base, opts),
      rust_lang(base, opts)
    ]
    |> Enum.reject(&is_nil/1)
  end

  defp elixir_lang do
    %{
      name: "elixir",
      available: true,
      cmd: fn spec, mode, input, output ->
        {"mix", ["run", script_path("elixir/bench.exs"), spec, mode, input, output]}
      end
    }
  end

  defp python_lang(base, opts) do
    cond do
      opts[:skip_python] ->
        nil

      System.find_executable("python3") == nil ->
        nil

      not File.exists?(Path.join(base, "python/bench.py")) ->
        nil

      true ->
        %{
          name: "python",
          available: true,
          cmd: fn spec, mode, input, output ->
            {"python3", [Path.join(base, "python/bench.py"), spec, mode, input, output]}
          end
        }
    end
  end

  defp rust_lang(base, opts) do
    cond do
      opts[:skip_rust] ->
        nil

      System.find_executable("cargo") == nil ->
        nil

      not File.exists?(Path.join(base, "rust/Cargo.toml")) ->
        nil

      true ->
        binary = build_rust_binary(base)

        if binary do
          %{
            name: "rust",
            available: true,
            cmd: fn spec, mode, input, output ->
              {binary, [spec, mode, input, output]}
            end
          }
        end
    end
  end

  defp build_rust_binary(base) do
    manifest = Path.join(base, "rust/Cargo.toml")
    IO.puts("Building Rust benchmark (release mode)...")

    case System.cmd("cargo", ["build", "--release", "--manifest-path", manifest],
           stderr_to_stdout: true
         ) do
      {_out, 0} ->
        Path.join(base, "rust/target/release/bench")

      {out, code} ->
        IO.puts(:stderr, "Rust build failed (exit #{code}):")
        IO.puts(:stderr, out)
        nil
    end
  end

  defp script_path(rel) do
    Path.join([File.cwd!(), "bench/comparative", rel])
  end

  defp prepare_run_dir(base) do
    stamp = DateTime.utc_now() |> Calendar.strftime("%Y-%m-%d-%H%M%S")
    dir = Path.join([base, "results", "run-#{stamp}"])
    File.mkdir_p!(dir)
    dir
  end

  defp run_phases(langs, spec_path, run_dir) do
    lang_map = Map.new(langs, &{&1.name, &1})
    available = Map.keys(lang_map)

    phase1 =
      for name <- available do
        out = Path.join(run_dir, "phase1-#{name}.xlsx")
        lang = Map.fetch!(lang_map, name)
        result = execute(lang, spec_path, "create", "-", out)
        Map.merge(result, %{"phase" => 1, "lang" => name, "output" => out})
      end

    phase2 = cross_phase(2, phase1, rotation_a(available), lang_map, spec_path, run_dir)
    phase3 = cross_phase(3, phase2, rotation_a(available), lang_map, spec_path, run_dir)
    phase4 = cross_phase(4, phase3, rotation_b(available), lang_map, spec_path, run_dir)

    phase1 ++ phase2 ++ phase3 ++ phase4
  end

  defp rotation_a(available) do
    case available do
      ["elixir", "python", "rust"] ->
        %{"rust" => "elixir", "python" => "rust", "elixir" => "python"}

      ["elixir", "python"] ->
        %{"python" => "elixir", "elixir" => "python"}

      ["elixir", "rust"] ->
        %{"rust" => "elixir", "elixir" => "rust"}

      ["elixir"] ->
        %{"elixir" => "elixir"}

      _ ->
        rotate_map(available, 1)
    end
  end

  defp rotation_b(available) do
    case available do
      ["elixir", "python", "rust"] ->
        %{"python" => "elixir", "elixir" => "rust", "rust" => "python"}

      ["elixir", "python"] ->
        %{"elixir" => "python", "python" => "elixir"}

      ["elixir", "rust"] ->
        %{"elixir" => "rust", "rust" => "elixir"}

      ["elixir"] ->
        %{"elixir" => "elixir"}

      _ ->
        rotate_map(available, 2)
    end
  end

  defp rotate_map(list, offset) do
    n = length(list)

    list
    |> Enum.with_index()
    |> Map.new(fn {name, i} -> {name, Enum.at(list, rem(i + offset, n))} end)
  end

  defp cross_phase(phase_num, prev_results, rotation, lang_map, spec_path, run_dir) do
    prev_by_lang = Map.new(prev_results, &{&1["lang"], &1["output"]})

    for {reader_name, writer_name} <- rotation do
      input = Map.fetch!(prev_by_lang, writer_name)
      out = Path.join(run_dir, "phase#{phase_num}-#{reader_name}-reads-#{writer_name}.xlsx")
      lang = Map.fetch!(lang_map, reader_name)
      result = execute(lang, spec_path, "edit", input, out)

      Map.merge(result, %{
        "phase" => phase_num,
        "lang" => reader_name,
        "reads" => writer_name,
        "output" => out
      })
    end
  end

  defp execute(lang, spec_path, mode, input, output) do
    {cmd, args} = lang.cmd.(spec_path, mode, input, output)

    {runner, runner_args} = wrap_with_time(cmd, args)

    IO.write("  #{lang.name} #{mode} #{Path.basename(output)}... ")

    {out, code} = System.cmd(runner, runner_args, stderr_to_stdout: true)

    {json_line, time_output} = split_time_output(out)

    result =
      case :json.decode(json_line) do
        decoded when is_map(decoded) ->
          decoded

        _ ->
          IO.puts("")
          IO.puts(:stderr, "Unexpected output:\n#{out}")
          %{"wall_ms" => nil, "cells_written" => 0, "cells_cleared" => 0}
      end

    rss_kb = parse_rss(time_output)
    status = if code == 0, do: "ok", else: "FAIL (#{code})"
    IO.puts("#{result["wall_ms"]} ms, #{(rss_kb && "#{rss_kb} KB") || "?"} peak — #{status}")

    Map.merge(result, %{"rss_kb" => rss_kb, "exit" => code})
  end

  defp wrap_with_time(cmd, args) do
    case :os.type() do
      {:unix, :darwin} ->
        {"/usr/bin/time", ["-l", cmd] ++ args}

      {:unix, _} ->
        {"/usr/bin/time", ["-v", cmd] ++ args}

      _ ->
        {cmd, args}
    end
  end

  defp split_time_output(out) do
    lines = String.split(out, "\n")

    {json_lines, rest} =
      Enum.split_with(lines, fn line ->
        String.starts_with?(line, "{")
      end)

    json = json_lines |> List.last() || "{}"
    {json, Enum.join(rest, "\n")}
  end

  defp parse_rss(text) do
    # macOS: "     12345678  maximum resident set size" (bytes)
    # Linux: "Maximum resident set size (kbytes): 12345"
    cond do
      match = Regex.run(~r/Maximum resident set size \(kbytes\):\s*(\d+)/, text) ->
        match |> Enum.at(1) |> String.to_integer()

      match = Regex.run(~r/^\s*(\d+)\s+maximum resident set size/m, text) ->
        bytes = match |> Enum.at(1) |> String.to_integer()
        div(bytes, 1024)

      true ->
        nil
    end
  end

  defp write_results(results, run_dir) do
    path = Path.join(run_dir, "results.json")
    File.write!(path, :json.encode(results))
  end

  defp generate_report(results, run_dir, _langs) do
    File.write!(Path.join(run_dir, "report.md"), Report.markdown(results))
    File.write!(Path.join(run_dir, "report.html"), Report.html(results))
  end
end
