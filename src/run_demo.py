from pathlib import Path
from billing_core import stage1_compute_billing_detail, stage2_generate_report

root = Path(__file__).resolve().parents[1]
out = root / "outputs"
out.mkdir(exist_ok=True)
detail = stage1_compute_billing_detail(str(root / "examples/input.xlsx"))
stage2_generate_report(detail, str(out / "billing_report.xlsx"))
detail.to_csv(out / "billing_detail.csv", index=False, encoding="utf-8-sig")
print(out / "billing_report.xlsx")
