# Pairwise Test Suite Generator

A Python tool that generates **pairwise (2-way) combinatorial test suites** from parameter definitions in Excel. Uses a hybrid strategy: a fast greedy algorithm for an initial solution, with an optional exact solver ([OR-Tools CP-SAT](https://developers.google.com/optimization)) to find the provably minimal suite.

## Features

- **Greedy algorithm** with smart heuristics (hardest-pair seeding, most-constrained-first fill)
- **Exact CP-SAT solver** finds the true minimum when the search space is tractable
- **Excel I/O** — reads parameters from `.xlsx`, writes the test suite + summary
- **Built-in verification** — every output suite is checked for full pairwise coverage
- **Configurable** — control Cartesian threshold, time limits, and exact/greedy mode


https://github.com/user-attachments/assets/c1dfcf74-df8d-4aa8-a200-e9984363aaed


## Installation

```bash
pip install openpyxl ortools
```

## Quick Start

```bash
python pairwise.py --input params.xlsx --output suite.xlsx
```

### Input Format

An Excel file where **Row 1** contains parameter names and **Rows 2+** list possible values per column. Empty cells are ignored. Parameters can have different numbers of values.

| Browser | OS      | Payment |
|---------|---------|---------|
| Chrome  | Windows | Card    |
| Firefox | Linux   | PayPal  |
| Safari  | macOS   |         |

### Output Format

Two sheets in the output Excel file:

- **Suite** — the generated test cases (same column headers as input)
- **Summary** — metrics: LB, UB, FINAL_SIZE, OPTIMAL_PROVEN, T, T_MAX, etc.

## CLI Options

| Flag | Description | Default |
|------|-------------|---------|
| `--input` | Path to input Excel file | *(required)* |
| `--output` | Path to output Excel file | *(required)* |
| `--sheet` | Sheet name to read | first sheet |
| `--seed` | Random seed for solver | `0` |
| `--T_MAX` | Cartesian size threshold for exact phase | `50000` |
| `--exact_time_limit` | Time limit (seconds) for exact phase only | unlimited |
| `--no-exact` | Skip exact phase, use greedy only | off |
| `--verbose` | Print progress logs | off |

### Examples

```bash
# Greedy only (fast)
python pairwise.py --input params.xlsx --output suite.xlsx --no-exact

# With exact solver, 60s time limit
python pairwise.py --input params.xlsx --output suite.xlsx --exact_time_limit 60

# Verbose output
python pairwise.py --input params.xlsx --output suite.xlsx --verbose
```

## Algorithm

1. **Parse** parameter names and values from Excel
2. **Compute bounds** — lower bound LB = max(vᵢ × vⱼ), Cartesian size T = ∏vᵢ
3. **Greedy phase** — deterministic algorithm seeded by the hardest uncovered pair, filling parameters in most-constrained-first order
4. **Exact phase** (optional) — if T ≤ T_MAX, builds a CP-SAT model over the full Cartesian product and solves `Minimize(∑xₜ)` subject to pairwise coverage constraints
5. **Verify** full pairwise coverage before writing output

## Benchmarks

Run the benchmark suite against standard configurations from the [literature](https://www.sciencedirect.com/science/article/pii/S0012365X0400130X#SEC5):

```bash
# Greedy only (all 16 benchmarks)
python benchmark.py

# With exact solver
python benchmark.py --exact --exact_time_limit 30

# Specific benchmarks
python benchmark.py --benchmarks "3^4" "3^13" "4^10"
```

### Sample Results (Greedy)

| Config | Greedy | IPO | TConfig | AETG | CTS |
|--------|--------|-----|---------|------|-----|
| 3⁴     | 11     | 10  | 9       | 9    | 9   |
| 3¹³    | 20     | 20  | 15      | 15   | 15  |
| 2¹⁰⁰   | 36     | 15  | 14      | —    | 10  |
| 4¹⁰    | 33     | 31  | 28      | —    | 28  |
| 4¹⁰⁰   | 91     | 52  | 43      | —    | 43  |

## Dependencies

- [openpyxl](https://openpyxl.readthedocs.io/) — Excel read/write
- [OR-Tools](https://developers.google.com/optimization) — CP-SAT constraint solver (exact phase)
