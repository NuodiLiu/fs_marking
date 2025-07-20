# manually import COM constants after makepy
import win32com.client.gencache

from core.result_writers.excel_summary_writer import ExcelSummaryWriter
from core.result_writers.feedback_writer import FeedbackWriter
win32com.client.gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 7)
win32com.client.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8)

from config.test1_config import test1_config
from core.pipeline import run_batch
from core.utils.utils import validate_config
from core.result_writer import ResultWriter
from core.result_writers.stdout_writer import StdoutWriter

def main():
    config = test1_config
    validate_config(config)

    excel_writer = ExcelSummaryWriter()
    writer = ResultWriter([StdoutWriter(), FeedbackWriter(), excel_writer])
    
    run_batch(config, writer)
    excel_writer.save()

if __name__ == "__main__":
    main()
    
