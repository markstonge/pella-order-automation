from __future__ import annotations

from pathlib import Path

from pella_order_automation.extractor import build_job
from pella_order_automation.workbook_writer import write_workbook


SAMPLE_DIR = Path("/Users/mark/Downloads/aipellaordersprocessingproject")
OUTPUT_DIR = Path("outputs")


def main() -> int:
    OUTPUT_DIR.mkdir(exist_ok=True)
    cases = [
        (
            SAMPLE_DIR / "Purchase Order Example 1.csv",
            SAMPLE_DIR / "Work Order Example 1.xlsx",
            OUTPUT_DIR / "smoke-example-1.xlsx",
        ),
        (
            SAMPLE_DIR / "Purchase Order Example 2.csv",
            SAMPLE_DIR / "Work Order Example 2.xlsx",
            OUTPUT_DIR / "smoke-example-2.xlsx",
        ),
    ]
    for po_path, work_order_path, output_path in cases:
        job = build_job(po_path, work_order_path)
        write_workbook(job, output_path)
        warnings = [*job.warnings, *(warning for line in job.jamb_lines for warning in line.warnings)]
        print(f"{output_path}: {len(job.jamb_lines)} jamb line(s), {len(warnings)} warning(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
