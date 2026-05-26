"""P5.6: Smoke check checkpoint path is portable.

Chạy: python tests/smoke_checkpoint.py
"""
from __future__ import annotations

import os
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from tests._helpers import setup_env, Report


cleanup = setup_env()
try:
    from word_backend import (
        _checkpoint_path, checkpoint_save, checkpoint_load, checkpoint_clear,
    )

    report = Report("smoke_checkpoint")

    tmpdir = tempfile.gettempdir()
    path1 = _checkpoint_path(b"file1-content", "Vietnamese")
    path2 = _checkpoint_path(b"file1-content", "English")
    path3 = _checkpoint_path(b"different-bytes", "Vietnamese")

    report.check("path under tempfile.gettempdir()",
                 path1.startswith(tmpdir),
                 f"path={path1}, tmpdir={tmpdir}")
    report.check("path does NOT hard-code /tmp on non-Unix systems",
                 "/tmp/tr_ckpt_" not in path1 or tmpdir == "/tmp",
                 f"OK on this platform; path={path1}")
    report.check("different target_lang → different path", path1 != path2)
    report.check("different docx bytes → different path", path1 != path3)

    # Save → load → clear round-trip
    translations = {"p0": "Xin chào", "p1": "Thế giới"}
    checkpoint_save(b"file1-content", "Vietnamese", translations)
    loaded = checkpoint_load(b"file1-content", "Vietnamese")
    report.check("checkpoint round-trip", loaded == translations,
                 f"loaded={loaded}")

    # Load with mismatched bytes returns None
    other = checkpoint_load(b"different-bytes", "Vietnamese")
    report.check("checkpoint isolated by docx hash", other is None)

    # Clear
    checkpoint_clear(b"file1-content", "Vietnamese")
    loaded_after_clear = checkpoint_load(b"file1-content", "Vietnamese")
    report.check("checkpoint cleared", loaded_after_clear is None)

finally:
    cleanup()


sys.exit(report.summary())
