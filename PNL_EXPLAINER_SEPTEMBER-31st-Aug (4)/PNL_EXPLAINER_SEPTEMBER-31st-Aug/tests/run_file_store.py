#!/usr/bin/env python3
"""Execute tests/test_file_store.bas against the real modStore source.

Everything under test here is pure: it decides what will be written, without
writing anything.  The position keys, the run fingerprint, the CSV quoting and
the folder-name ordering are where a store goes wrong QUIETLY - a name with a
comma shifts one row's columns, an unpadded month serves the wrong day - and
none of them needs a filesystem to be wrong.

Not covered here, because it needs Windows: ADODB.Stream, MkDir and Dir$.
Those fail loudly; these would not.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import modules                                            # noqa: E402
from vba_run import BasicRunner                           # noqa: E402

PURE_FROM_MODSTORE = [
    # position identity
    "StoreBondPositionKey",
    "StoreSwapPositionKey",
    "StoreFuturePositionKey",
    "StoreJoinKey",
    "StoreClean",
    # run identity
    "StoreHash32",
    "StoreXor32",
    "StoreMul32",
    "StoreHex32",
    "StoreFingerprintOfKeys",
    # what gets written
    "StoreNumberText",
    "StoreLooksNumeric",
    "StoreTypedValue",
    "StoreTypeArray",
    "StoreCsvField",
    "StoreCsvText",
    "StoreCsvParse",
    "StoreDateText",
    "StoreRunFolderName",
    "StorePad",
]


def main() -> int:
    r = BasicRunner()
    r.add_procedures(modules.path_of("modStore"), PURE_FROM_MODSTORE)
    r.add_source((ROOT / "tests" / "test_file_store.bas").read_text())

    out = r.run("RunTests")
    print(out)
    return 0 if "failed=0" in out else 1


if __name__ == "__main__":
    raise SystemExit(main())
