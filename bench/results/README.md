# ExVEx Performance Benchmark Report

Benchmarks captured with Benchee (5 s timed runs, 2 s warmup, 1 s memory
measurement). Run with:

```bash
BENCH_LABEL=baseline mix run bench/put_cell_benchmark.exs
BENCH_LABEL=after BENCH_LOAD=bench/results/baseline.benchee mix run bench/put_cell_benchmark.exs
```

## Environment

- Apple M3 Max, macOS, 48 GB RAM, 16 cores
- Elixir 1.18.3 / Erlang 27.2, JIT enabled
- Source fixture: `test/fixtures/cells.xlsx`

## Summary — what changed

Two targeted fixes land in the same commit:

1. **Sheet tree cache on `%ExVEx.Workbook{}`.** The old `put_cell` /
   `get_cell` / `merge_cells` / `unmerge_cells` paths all parsed the full
   worksheet XML (`Saxy.SimpleForm.parse_string`) and re-emitted it
   (`Saxy.encode!`) on every call. For bulk writes this went quadratic.
   Now each sheet's parsed tree is cached on first access and re-used
   across subsequent calls; `save/2` re-serializes only dirty sheet
   trees once at flush time.

2. **`ExVEx.OOXML.SharedStrings` switched from tuple to two maps**
   (`by_index`, `by_string`). `intern/2` is now O(1) instead of O(N) per
   call (was `Tuple.insert_at` which copied the N-element tuple each
   time). `get/2` stays O(1).

Public API is unchanged. Round-trip fidelity tests still pass.

## Results

| scenario | baseline | after | speedup |
| --- | ---: | ---: | ---: |
| 500 `put_cell` + save | 141.89 ms / 348 MB | **6.85 ms / 5.86 MB** | **~20× faster, 59× less memory** |
| 1000 unique string interns + save | 761.49 ms / 1.8 GB | **35.64 ms / 16.2 MB** | **~21× faster, 111× less memory** |
| 500 `get_cell` | 16.77 ms / 28 MB | 16.94 ms / 28 MB | unchanged |

Read-only `get_cell` is unchanged: the sheet tree is now cached on first
access, but the resolution path still calls `cells_from_tree/1` on every
lookup. Walking the tree is cheap compared to re-parsing the XML (which
is the win for `put_cell`) but could be further improved by memoizing the
extracted `%{coord => cell}` map. That's a follow-up optimization, not
blocking this release.

## Detailed Benchee output

### Baseline (`bench/results/baseline.txt`)

```
Name                                        ips        average  deviation         median         99th %
500 get_cell                              59.63       16.77 ms     ±1.40%       16.74 ms       17.63 ms
500 put_cell + save                        7.05      141.89 ms     ±3.07%      143.42 ms      147.38 ms
1000 unique string interns + save          1.31      761.49 ms     ±0.86%      759.53 ms      770.80 ms

Memory usage statistics:
500 get_cell                             28.06 MB
500 put_cell + save                     347.97 MB
1000 unique string interns + save      1803.93 MB
```

### After (`bench/results/after.txt`)

```
Name                                        ips        average  deviation         median         99th %
500 put_cell + save                      146.07        6.85 ms    ±17.72%        6.73 ms        8.52 ms
500 get_cell                              59.05       16.94 ms     ±1.79%       16.86 ms       17.96 ms
1000 unique string interns + save         28.06       35.64 ms     ±4.96%       35.38 ms       46.06 ms

Memory usage statistics:
500 put_cell + save                       5.86 MB
500 get_cell                             28.16 MB
1000 unique string interns + save        16.23 MB
```

## Reproducing

```bash
# Baseline (before the optimization commit)
git checkout <pre-perf-commit>
BENCH_LABEL=baseline mix run bench/put_cell_benchmark.exs

# After (current main)
git checkout main
BENCH_LABEL=after BENCH_LOAD=bench/results/baseline.benchee mix run bench/put_cell_benchmark.exs
```

The `BENCH_LOAD` variable makes Benchee emit a side-by-side comparison.
