# manually import COM constants after makepy
import win32com.client.gencache
import sys
import cProfile
import pstats
import io

from core.result_writers.excel_summary_writer import ExcelSummaryWriter
from core.result_writers.feedback_writer import FeedbackWriter

win32com.client.gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 7)
win32com.client.gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8)

from core.config_loader import ConfigLoader
from core.pipeline import run_batch
from core.result_writer import ResultWriter
from core.result_writers.stdout_writer import StdoutWriter

def main():
    # Check for --profile and --benchmark flags
    enable_profile = "--profile" in sys.argv
    if enable_profile:
        sys.argv.remove("--profile")
    
    enable_benchmark = "--benchmark" in sys.argv
    if enable_benchmark:
        sys.argv.remove("--benchmark")
    
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
    import time
    if enable_benchmark:
        t0 = time.perf_counter()
    excel_writer = ExcelSummaryWriter()
    if enable_benchmark:
        excel_load_time = time.perf_counter() - t0
    
    writer = ResultWriter([StdoutWriter(), FeedbackWriter(), excel_writer])

    # Run batch evaluation with profiling/benchmark
    if enable_profile:
        print("🔬 Profiling enabled...\n")
        profiler = cProfile.Profile()
        profiler.enable()
        
        run_batch(exam_config, writer, benchmark=enable_benchmark)
        
        profiler.disable()
        
        # Print profiling results
        print("\n" + "="*80)
        print("📊 PROFILING RESULTS")
        print("="*80)
        
        s = io.StringIO()
        stats = pstats.Stats(profiler, stream=s)
        
        # Sort by cumulative time
        stats.sort_stats('cumulative')
        stats.print_stats(30)  # Top 30 functions
        
        print(s.getvalue())
        
        # Save detailed profile to file
        stats.dump_stats('profile_results.prof')
        print("\n💾 Detailed profile saved to: profile_results.prof")
        print("📝 View with: python -m pstats profile_results.prof")
    else:
        run_batch(exam_config, writer, benchmark=enable_benchmark)
    
    # Save results
    if enable_benchmark:
        t0 = time.perf_counter()
    excel_writer.save()
    if enable_benchmark:
        excel_save_time = time.perf_counter() - t0
        print(f"\n📊 Excel Operations:")
        print(f"  • Load template: {excel_load_time*1000:.2f} ms")
        print(f"  • Save file: {excel_save_time*1000:.2f} ms")
    
    print(f"\n✅ Finished marking for {exam_config.name}")

if __name__ == "__main__":
    main()
