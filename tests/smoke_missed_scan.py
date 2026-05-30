from __future__ import annotations

import sys

from _helpers import Report, setup_env


def block(block_id: str, text: str, role: str = "paragraph") -> dict:
    return {
        "id": block_id,
        "text": text,
        "text_tagged": text,
        "has_format": False,
        "role": role,
        "para_idx": int(block_id[1:]) if block_id.startswith("p") and block_id[1:].isdigit() else -1,
        "table_cell": None,
    }


def main() -> int:
    cleanup = setup_env()
    try:
        from word_backend import find_missed, has_source_language_residue

        report = Report("smoke_missed_scan")

        b0 = block("p0", "The landing door shall be inspected.")
        report.check(
            "exact original is missed",
            find_missed([b0], {"p0": b0["text"]}, "English", "Vietnamese") == [b0],
        )

        report.check(
            "english sentence residue detected",
            has_source_language_residue(
                "Cua tang shall be inspected before operation.",
                "English",
                "Vietnamese",
            ),
        )

        report.check(
            "technical english phrase residue detected",
            bool(find_missed(
                [b0],
                {"p0": "Kiem tra landing door truoc khi van hanh."},
                "English",
                "Vietnamese",
            )),
        )

        b1 = block("p1", "Applicable standard")
        report.check(
            "standards and model codes allowed",
            not find_missed(
                [b1],
                {"p1": "Tieu chuan ap dung: EN 81-20, ISO 22201, Schindler 3300."},
                "English",
                "Vietnamese",
            ),
        )

        report.check(
            "output block scan catches apply miss",
            bool(find_missed(
                [b0],
                {"p0": "Cua tang phai duoc kiem tra."},
                "English",
                "Vietnamese",
                output_blocks=[block("p0", "The landing door shall be inspected.")],
            )),
        )

        report.check(
            "vietnamese residue detected in english target",
            has_source_language_residue(
                "The cabin chua duoc kiem tra.",
                "Vietnamese",
                "English",
            ),
        )

        report.check(
            "clean english target accepted",
            not has_source_language_residue(
                "The cabin has not been inspected.",
                "Vietnamese",
                "English",
            ),
        )

        return report.summary()
    finally:
        cleanup()


if __name__ == "__main__":
    sys.exit(main())
