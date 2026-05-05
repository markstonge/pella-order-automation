from pathlib import Path

from pella_order_automation.extractor import build_job
from pella_order_automation.po_parser import parse_po_csv
from pella_order_automation.work_order_parser import parse_work_order_xlsx


SAMPLE_DIR = Path("/Users/mark/Downloads/aipellaordersprocessingproject")


def test_parse_purchase_order_example_1():
    lines, warnings = parse_po_csv(SAMPLE_DIR / "Purchase Order Example 1.csv")
    assert not warnings
    assert len(lines) == 4
    assert lines[0].pella_po_number == "408C222504"
    assert lines[0].pella_xref == "4083AM488A-005"


def test_parse_work_order_example_1():
    work_order, warnings = parse_work_order_xlsx(SAMPLE_DIR / "Work Order Example 1.xlsx")
    assert not warnings
    assert work_order.order_number == "4083AM488A"
    assert 5 in work_order.items


def test_build_job_example_1():
    job = build_job(
        SAMPLE_DIR / "Purchase Order Example 1.csv",
        SAMPLE_DIR / "Work Order Example 1.xlsx",
    )
    assert job.order_number == "4083AM488A"
    assert len(job.jamb_lines) == 4
    assert job.jamb_lines[0].jamb_depth == 3.5625
