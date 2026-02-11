# manually import COM constants after makepy
import win32com.client.gencache
import sys

from core.result_writers.excel_summary_writer import ExcelSummaryWriter
from core.result_writers.feedback_writer import FeedbackWriter

win32com.client.gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 7)
win32com.client.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8)

from core.config_loader import ConfigLoader
from core.pipeline import run_batch
from core.result_writer import ResultWriter
from core.result_writers.stdout_writer import StdoutWriter

def main():
    # Select exam configuration from command line or default
    if len(sys.argv) > 1:
        exam_code = sys.argv[1]
    else:
        exam_code = "test1"  # Default
    
    print(f"📋 Loading configuration: {exam_code}")
    
    # Load configuration
    try:
        exam_config = ConfigLoader.load_from_module(f"config.{exam_code}_config")
    except (ImportError, ModuleNotFoundError):
        print(f"❌ Configuration not found: config.{exam_code}_config")
        print(f"Available configs: test1_config, test2_config")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Failed to load configuration: {e}")
        sys.exit(1)
    
    print(f"✅ Loaded: {exam_config.name}")
    print(f"📏 Total rules: {len(exam_config.rules)}")
    print(f"🎯 Total marks: {exam_config.total_marks}")
    
    # Create writers (dependency injection)
    excel_writer = ExcelSummaryWriter()
    writer = ResultWriter([StdoutWriter(), FeedbackWriter(), excel_writer])

    # Run batch evaluation
    run_batch(exam_config, writer)
    
    # Save results
    excel_writer.save()
    
    print(f"\n✅ Finished marking for {exam_config.name}")

if __name__ == "__main__":
    main()
