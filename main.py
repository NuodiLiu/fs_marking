# manually import COM constants after makepy
import win32com.client.gencache
win32com.client.gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 7)
win32com.client.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8)

from config.test1_config import test1_config
from core.pipeline import run_batch
from core.utils.utils import validate_config
from core.result_writer import ResultWriter
from core.result_writers.stdout_writer import StdoutWriter
from core.result_writers.excel_summary_writer import ExcelSummaryWriter

def main():
    config = test1_config
    validate_config(config)

    out = config.OUTPUT

    # 1) 自动收集所有规则行号
    rule_rows = {}
    for key, val in out.items():
        if not key.endswith("_row"):
            continue
        if key in ("zid_row", "total_row"):
            continue
        # key = "margin_row" → rule = "margin"
        rule = key[:-4].lower()
        rule_rows[rule] = val

    # 2) 取出其他两个
    zid_row   = out["zid_row"]
    total_row = out["total_row"]

    excel_writer = ExcelSummaryWriter(
        template_path = out["template_path"],
        output_path   = out["output_path"],
        sheet_name    = out["sheet_name"],
        zid_row       = zid_row,
        rule_rows     = rule_rows,
        total_row     = total_row,
    )

    writer = ResultWriter([StdoutWriter(), excel_writer])
    run_batch(config, writer)
    excel_writer.save()

    print("🔍 当前启用的规则及其 name：")
    for r in config.RULES:
        print(f"  类 {r.__class__.__name__} → name = {r.name}")

if __name__ == "__main__":
    main()
    
