defmodule ExVEx.Bench.Report do
  @moduledoc false

  def markdown(results) do
    by_phase = Enum.group_by(results, & &1["phase"])

    header = """
    # Comparative xlsx Benchmark — Run #{DateTime.utc_now() |> Calendar.strftime("%Y-%m-%d %H:%M UTC")}

    Per-phase wall time and peak resident-set-size for each participating
    runtime. Lower is better for both columns. RSS is measured externally via
    `/usr/bin/time`; a `—` means the host OS didn't report it in a parseable
    form.

    """

    sections =
      for phase <- Enum.sort(Map.keys(by_phase)) do
        rows = Map.fetch!(by_phase, phase)

        """
        ## Phase #{phase}

        #{phase_description(phase)}

        | Runtime | Mode | Reads | Wall time (ms) | Peak RSS (MB) | Cells written | Cells cleared |
        |---------|------|-------|---------------:|--------------:|--------------:|--------------:|
        #{Enum.map_join(rows, "\n", &markdown_row/1)}
        """
      end

    header <> Enum.join(sections, "\n")
  end

  defp markdown_row(r) do
    mode = r["mode"]
    reads = r["reads"] || "—"
    wall = format_number(r["wall_ms"])

    rss =
      (r["rss_kb"] && Float.round(r["rss_kb"] / 1024, 1) |> :erlang.float_to_binary(decimals: 1)) ||
        "—"

    cw = r["cells_written"]
    cc = r["cells_cleared"]
    "| #{r["lang"]} | #{mode} | #{reads} | #{wall} | #{rss} | #{cw} | #{cc} |"
  end

  defp format_number(nil), do: "—"
  defp format_number(n) when is_integer(n), do: Integer.to_string(n)

  defp format_number(n) when is_float(n),
    do: :erlang.float_to_binary(n, decimals: 1)

  defp phase_description(1),
    do: "Each runtime creates its own template from scratch."

  defp phase_description(2),
    do: "Rotation A — each runtime reads and edits another runtime's Phase 1 output."

  defp phase_description(3),
    do: "Same rotation as Phase 2 with Phase 2 outputs as inputs (double round-trip)."

  defp phase_description(4),
    do: "Rotation B — different pairings to cover the remaining interop combinations."

  def html(results) do
    by_phase = Enum.group_by(results, & &1["phase"])

    payload =
      results
      |> Enum.map(
        &Map.take(&1, ~w(phase lang reads mode wall_ms rss_kb cells_written cells_cleared))
      )

    json = :json.encode(payload)

    """
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <title>ExVEx comparative benchmark</title>
      <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 1000px; margin: 2rem auto; padding: 0 1rem; color: #1f2328; }
        h1 { margin-bottom: 0.25rem; }
        h2 { margin-top: 2rem; }
        .meta { color: #656d76; font-size: 0.9rem; margin-bottom: 2rem; }
        .chart-row { display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin: 1rem 0; }
        .chart-row canvas { background: #f6f8fa; border-radius: 6px; padding: 0.5rem; }
        table { border-collapse: collapse; width: 100%; font-size: 0.9rem; margin: 0.5rem 0; }
        th, td { padding: 0.4rem 0.6rem; border-bottom: 1px solid #d0d7de; text-align: left; }
        th { background: #f6f8fa; }
        td.num { text-align: right; font-variant-numeric: tabular-nums; }
      </style>
    </head>
    <body>
      <h1>ExVEx comparative benchmark</h1>
      <div class="meta">#{length(by_phase |> Map.keys())} phases, #{length(results)} operations. Run at #{DateTime.utc_now() |> Calendar.strftime("%Y-%m-%d %H:%M UTC")}.</div>

      #{for phase <- Enum.sort(Map.keys(by_phase)), do: html_phase(phase, Map.fetch!(by_phase, phase))}

      <script>
      const results = #{json};
      const phases = [...new Set(results.map(r => r.phase))].sort();
      const langs = [...new Set(results.map(r => r.lang))].sort();
      const colors = { elixir: '#6e40c9', python: '#0969da', rust: '#cf222e' };

      function chartFor(canvasId, phase, metric, label) {
        const rows = results.filter(r => r.phase === phase);
        const data = {
          labels: rows.map(r => r.reads ? `${r.lang} ← ${r.reads}` : r.lang),
          datasets: [{
            label: label,
            data: rows.map(r => r[metric]),
            backgroundColor: rows.map(r => colors[r.lang] || '#888')
          }]
        };
        new Chart(document.getElementById(canvasId), {
          type: 'bar',
          data: data,
          options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, title: { display: true, text: label } } } }
        });
      }

      phases.forEach(p => {
        chartFor(`wall-${p}`, p, 'wall_ms', 'Wall time (ms)');
        chartFor(`rss-${p}`, p, 'rss_kb', 'Peak RSS (KB)');
      });
      </script>
    </body>
    </html>
    """
  end

  defp html_phase(phase, rows) do
    """
    <h2>Phase #{phase}</h2>
    <p>#{phase_description(phase)}</p>
    <div class="chart-row">
      <canvas id="wall-#{phase}" height="200"></canvas>
      <canvas id="rss-#{phase}" height="200"></canvas>
    </div>
    <table>
      <thead><tr><th>Runtime</th><th>Mode</th><th>Reads</th><th>Wall (ms)</th><th>RSS (MB)</th><th>Cells written</th><th>Cells cleared</th></tr></thead>
      <tbody>
      #{Enum.map_join(rows, "\n", &html_row/1)}
      </tbody>
    </table>
    """
  end

  defp html_row(r) do
    rss_mb =
      case r["rss_kb"] do
        nil -> "—"
        kb -> :erlang.float_to_binary(kb / 1024, decimals: 1)
      end

    "<tr><td>#{r["lang"]}</td><td>#{r["mode"]}</td><td>#{r["reads"] || "—"}</td>" <>
      "<td class=\"num\">#{format_number(r["wall_ms"])}</td>" <>
      "<td class=\"num\">#{rss_mb}</td>" <>
      "<td class=\"num\">#{r["cells_written"]}</td>" <>
      "<td class=\"num\">#{r["cells_cleared"]}</td></tr>"
  end
end
