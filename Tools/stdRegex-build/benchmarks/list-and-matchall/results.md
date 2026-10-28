# stdRegex before/after benchmark

Host: Word 16.0, 64-bit VBA.
Measured UTC: 2026-10-28T00:02:45.2888070Z.
Baseline: `c02d86fd7e5c99b648f9729d7a0c6160d492a642` (pure VBA engine, before list optimizations).

7 samples per operation/case; medians in milliseconds per call. QueryPerformanceCounter; compiled regexes reused; 2 warm-up calls; calibrated batches; rotated operation order; includes result allocation and release; output equality checked before timing.

ListArr uses two output columns. MatchAllArr materializes all numbered captures; MatchAll additionally creates dictionaries and named entries. Their contracts differ.

## No matches, 1,000 characters

- List before → List after: **2.689 → 2.696 ms**, 1.00× relative speed (0.3% more time).
- ListArr before → ListArr after: **2.692 → 2.694 ms**, 1.00× relative speed (0.1% more time).
- MatchAll → MatchAllArr: **2.807 → 3.998 ms**, 0.70× relative speed (42.4% more time).

## 1 match, 3 named captures

- List before → List after: **1.100 → 0.064 ms**, 17.19× relative speed (94.2% less time).
- ListArr before → ListArr after: **1.134 → 0.071 ms**, 15.90× relative speed (93.7% less time).
- MatchAll → MatchAllArr: **1.057 → 1.104 ms**, 0.96× relative speed (4.5% more time).

## 100 matches, 3 named captures

- List before → List after: **109.121 → 4.698 ms**, 23.23× relative speed (95.7% less time).
- ListArr before → ListArr after: **109.021 → 5.410 ms**, 20.15× relative speed (95.0% less time).
- MatchAll → MatchAllArr: **107.997 → 5.224 ms**, 20.67× relative speed (95.2% less time).

## 1,000 matches, 3 named captures

- List before → List after: **447.354 → 45.126 ms**, 9.91× relative speed (89.9% less time).
- ListArr before → ListArr after: **464.111 → 51.297 ms**, 9.05× relative speed (88.9% less time).
- MatchAll → MatchAllArr: **422.765 → 39.383 ms**, 10.73× relative speed (90.7% less time).

## 100 matches, no captures

- List before → List after: **44.856 → 6.552 ms**, 6.85× relative speed (85.4% less time).
- ListArr before → ListArr after: **44.780 → 7.047 ms**, 6.35× relative speed (84.3% less time).
- MatchAll → MatchAllArr: **42.952 → 6.052 ms**, 7.10× relative speed (85.9% less time).

## 100 sparse matches, 3 named captures

- List before → List after: **96.092 → 55.629 ms**, 1.73× relative speed (42.1% less time).
- ListArr before → ListArr after: **97.717 → 56.707 ms**, 1.72× relative speed (42.0% less time).
- MatchAll → MatchAllArr: **94.228 → 55.539 ms**, 1.70× relative speed (41.1% less time).

These measurements describe this host and these workloads. Sparse/no-match workloads spend more time in the unchanged search engine, so wrapper-only improvements may be small or negative.
