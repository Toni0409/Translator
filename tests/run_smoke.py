"""Run all smoke scripts in `tests/`. Trả về exit code != 0 nếu bất kỳ test fail.

Chạy:
    python tests/run_smoke.py
"""
from __future__ import annotations

import os
import subprocess
import sys


HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, ".."))

SCRIPTS = [
    "smoke_config.py",
    "smoke_docx.py",
    "smoke_domain.py",
    "smoke_checkpoint.py",
    "smoke_direction.py",
    "smoke_ocr.py",
]


def main() -> int:
    failures = []
    for script in SCRIPTS:
        path = os.path.join(HERE, script)
        if not os.path.exists(path):
            print(f"WARN Skip missing: {script}")
            continue
        print(f"\n=== {script} ===")
        r = subprocess.run([sys.executable, path], cwd=REPO)
        if r.returncode != 0:
            failures.append(script)

    print()
    if failures:
        print(f"FAIL {len(failures)} smoke script(s) failed: {', '.join(failures)}")
        return 1
    print(f"PASS All {len(SCRIPTS)} smoke scripts passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
