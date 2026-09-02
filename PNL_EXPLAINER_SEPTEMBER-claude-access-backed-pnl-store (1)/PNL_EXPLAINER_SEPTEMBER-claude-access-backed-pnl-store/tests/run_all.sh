#!/usr/bin/env bash
# Full test suite.  Run from the repository root.
set -uo pipefail
cd "$(dirname "$0")/.."

fail=0

echo "=== linter self-test (tests/test_lint_selftest.py) ==="
python3 tests/test_lint_selftest.py || fail=1

echo
echo "=== static analysis (tools/vba_lint.py) ==="
python3 tools/vba_lint.py || fail=1

echo
echo "=== line endings (tools/check_line_endings.py) ==="
python3 tools/check_line_endings.py || fail=1

echo
echo "=== layout / column contract (tools/check_layout.py) ==="
if [ -f tools/check_layout.py ]; then
    python3 tools/check_layout.py || fail=1
fi

echo
echo "=== Access schema agreement (tools/check_access_schema.py) ==="
python3 tools/check_access_schema.py || fail=1

echo
echo "=== Dashboard formulas (tools/check_dashboard.py) ==="
python3 tools/check_dashboard.py || fail=1

echo
echo "=== documentation diagrams (tools/check_diagrams.py) ==="
python3 tools/check_diagrams.py || fail=1

echo
echo "=== formula-string equivalence (LibreOffice Basic) ==="
python3 tests/run_formula_equivalence.py || fail=1

echo
echo "=== pull-to-par numerics (LibreOffice Basic) ==="
python3 tests/run_pull_to_par.py || fail=1

echo
echo "=== Bloomberg security fallback chain (LibreOffice Basic) ==="
python3 tests/run_bbg_fallback.py || fail=1

echo
echo "=== Access store keys and literals (LibreOffice Basic) ==="
python3 tests/run_access_store.py || fail=1

echo
if [ "$fail" -eq 0 ]; then
    echo "ALL SUITES PASSED"
else
    echo "SOME SUITES FAILED"
fi
exit "$fail"
