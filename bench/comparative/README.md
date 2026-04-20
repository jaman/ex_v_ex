# Comparative xlsx Benchmark

Cross-language time + memory benchmark across three xlsx implementations:

- **Elixir** — ExVEx (this library)
- **Python** — openpyxl
- **Rust** — umya-spreadsheet

See [`../../BENCHMARK.md`](../../../BENCHMARK.md) (in the project root) for
the methodology and four-phase rotation design.

## Running

From the `ex_v_ex/` project root:

```bash
mix bench.comparative
```

The orchestrator will:

1. Detect available runtimes (Python on `PATH`, `cargo` on `PATH`).
2. Run phase 1 (each impl creates its own template).
3. Run phases 2-4 (cross-read, edit, save).
4. Record per-op JSON to `bench/comparative/results/run-<timestamp>/`.
5. Emit `report.md` (markdown) and `report.html` (Chart.js).

## Implementations

Each implementation exposes a uniform CLI:

```
<runner> <spec_path> <mode> <input_or_-> <output>
```

- `<mode>` = `create` or `edit`
- `<input_or_->` = path to read (or `-` in create mode)

Each prints a single JSON line to stdout:

```
{"wall_ms": 42.5, "cells_written": 24000, "cells_cleared": 24000}
```

The orchestrator wraps each invocation with `/usr/bin/time` to capture peak
RSS externally. Both numbers end up in the report.

## Per-language entry points

| Lang | File | Invocation |
|------|------|------------|
| Elixir | `elixir/bench.exs` | `mix run bench/comparative/elixir/bench.exs ...` |
| Python | `python/bench.py` | `python3 bench/comparative/python/bench.py ...` |
| Rust | `rust/src/main.rs` | `cargo run --release --manifest-path bench/comparative/rust/Cargo.toml -- ...` |
