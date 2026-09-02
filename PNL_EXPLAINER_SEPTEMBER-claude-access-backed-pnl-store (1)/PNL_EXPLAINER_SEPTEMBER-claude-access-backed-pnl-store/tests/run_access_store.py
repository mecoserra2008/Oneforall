#!/usr/bin/env python3
"""Execute tests/test_access_store.bas against the real modAccess source.

Everything under test here is pure: it decides what the database will be told,
without touching one.  That is deliberate - the position keys, the run
fingerprint and the SQL literals are where a store goes wrong quietly, and none
of them needs Access to be wrong.

Not covered here, because it cannot be without a Windows machine and ACE: the
provider, the transaction, and the recordset append.  Those fail loudly; these
would not.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import modules                                            # noqa: E402
from vba_run import BasicRunner                           # noqa: E402

PURE_FROM_MODACCESS = [
    # position identity
    "AccBondPositionKey",
    "AccSwapPositionKey",
    "AccFuturePositionKey",
    "AccJoinKey",
    "AccClean",
    # run identity
    "AccHash32",
    "AccXor32",
    "AccMul32",
    "AccHex32",
    "AccFingerprintOfKeys",
    # statement text
    "AccSqlText",
    "AccSqlNum",
    "AccSqlDate",
    "AccPad",
    # DDL
    "AccCreateTableSql",
    "AccTableSpec",
    "AccTablePK",
    "AccSpecName",
    "AccSpecType",
]


def main() -> int:
    r = BasicRunner()
    r.add_procedures(modules.path_of("modAccess"), PURE_FROM_MODACCESS)
    r.add_source((ROOT / "tests" / "test_access_store.bas").read_text())

    out = r.run("RunTests")
    print(out)
    return 0 if "failed=0" in out else 1


if __name__ == "__main__":
    raise SystemExit(main())
