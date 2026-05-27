"""P5.2: Smoke check config import KHÔNG cần secrets.toml.

Chạy: python tests/smoke_config.py
"""
from __future__ import annotations

import os
import sys

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from tests._helpers import Report

report = Report("smoke_config")


# 1. Import config khi secrets.toml KHÔNG tồn tại
secrets = os.path.join(REPO_ROOT, ".streamlit", "secrets.toml")
backup = None
if os.path.exists(secrets):
    with open(secrets) as f:
        backup = f.read()
    os.remove(secrets)

try:
    for m in list(sys.modules):
        if m == "config":
            del sys.modules[m]
    import config
    report.check("import config without secrets.toml", True,
                 f"API_KEY={config.API_KEY!r}, APP_PASSWORD={config.APP_PASSWORD!r}")
    report.check("API_KEY empty fallback",      config.API_KEY == "")
    report.check("APP_PASSWORD empty fallback", config.APP_PASSWORD == "")
    report.check("LANGUAGES restricted to 2",   config.LANGUAGES == ["Tiếng Anh", "Tiếng Việt"])
    report.check("TRANSLATION_DIRECTIONS len 2", len(config.TRANSLATION_DIRECTIONS) == 2)
    report.check("PRICE_INPUT updated",         abs(config.PRICE_INPUT - 1.50) < 1e-9)
    report.check("PRICE_OUTPUT updated",        abs(config.PRICE_OUTPUT - 9.00) < 1e-9)
    report.check("media_only in NO_TRANSLATE_ROLES",
                 "media_only" in config.NO_TRANSLATE_ROLES)
except Exception as e:
    report.check(f"import config (raised {type(e).__name__})", False, str(e))
finally:
    if backup is not None:
        with open(secrets, "w") as f:
            f.write(backup)


# 2. calc_cost dùng pricing đúng (1.50 / 9.00)
try:
    from ui_common import calc_cost
    usd, vnd = calc_cost(100_000, 50_000)
    # 100K*1.50/1M + 50K*9.00/1M = 0.15 + 0.45 = 0.60
    report.check("calc_cost USD",
                 abs(usd - 0.60) < 1e-6,
                 f"got ${usd:.4f}, expected $0.6000")
    report.check("calc_cost VND",
                 abs(vnd - 15_240) < 1.0,
                 f"got {vnd:,.2f}, expected 15,240.00")
except Exception as e:
    report.check(f"calc_cost (raised {type(e).__name__})", False, str(e))


sys.exit(report.summary())
