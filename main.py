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
from core.pipeline import run_batch, run_batch_parallel
from core.result_writer import ResultWriter
from core.result_writers.stdout_writer import StdoutWriter
import multiprocessing

def main():
    # Check for --profile, --benchmark, and --parallel flags
    enable_profile = "--profile" in sys.argv
    if enable_profile:
        sys.argv.remove("--profile")
    
    enable_benchmark = "--benchmark" in sys.argv
    if enable_benchmark:
        sys.argv.remove("--benchmark")
    
    enable_parallel = "--parallel" in sys.argv
    if enable_parallel:
        sys.argv.remove("--parallel")
    
    # Check for -j flag (number of workers)
    num_workers = None
    if "-j" in sys.argv:
        idx = sys.argv.index("-j")
        if idx + 1 < len(sys.argv):
            try:
                num_workers = int(sys.argv[idx + 1])
                sys.argv.pop(idx + 1)  # Remove number
                sys.argv.pop(idx)  # Remove -j
            except ValueError:
                print("⚠️ Invalid -j value, using default")
    
    # Select exam configuration from command line or default
    if len(sys.argv) > 1:
        exam_code = sys.argv[1]
    else:
        exam_code = "test1"  # Default
    
    # Load configuration
    try:
        exam_config = ConfigLoader.load_from_module(f"config.{exam_code}_config")
    except (ImportError, ModuleNotFoundError):
        print(f"Configuration not found: config.{exam_code}_config")
        print(f"Available configs: test1_config, test2_config, exam_feb11_config")
        sys.exit(1)
    except Exception as e:
        print(f"Failed to load configuration: {e}")
        sys.exit(1)
    
    # Check for conflicting options
    if enable_parallel and enable_profile:
        print("Warning: --profile not effective with --parallel (only tracks main process)")
        enable_profile = False
    
    # Create writers (dependency injection)
    import time
    if enable_benchmark:
        t0 = time.perf_counter()
    excel_writer = ExcelSummaryWriter()
    if enable_benchmark:
        excel_load_time = time.perf_counter() - t0
    
    writer = ResultWriter([StdoutWriter(), FeedbackWriter(), excel_writer])

    # Run batch evaluation
    if enable_parallel:
        # Parallel mode
        if enable_profile:
            profiler = cProfile.Profile()
            profiler.enable()
            
            run_batch_parallel(exam_config, writer, benchmark=enable_benchmark, num_workers=num_workers)
            
            profiler.disable()
            
            # Print profiling results
            print("\n" + "="*80)
            print("PROFILING RESULTS")
            print("="*80)
            
            s = io.StringIO()
            stats = pstats.Stats(profiler, stream=s)
            stats.sort_stats('cumulative')
            stats.print_stats(30)
            print(s.getvalue())
            stats.dump_stats('profile_results.prof')
            print("\nProfile saved to: profile_results.prof")
            print("View with: python -m pstats profile_results.prof")
        else:
            run_batch_parallel(exam_config, writer, benchmark=enable_benchmark, num_workers=num_workers)
    
    elif enable_profile:
        # Serial mode with profiling
        profiler = cProfile.Profile()
        profiler.enable()
        
        run_batch(exam_config, writer, benchmark=enable_benchmark)
        
        profiler.disable()
        
        # Print profiling results
        print("\n" + "="*80)
        print("PROFILING RESULTS")
        print("="*80)
        
        s = io.StringIO()
        stats = pstats.Stats(profiler, stream=s)
        
        # Sort by cumulative time
        stats.sort_stats('cumulative')
        stats.print_stats(30)  # Top 30 functions
        
        print(s.getvalue())
        
        # Save detailed profile to file
        stats.dump_stats('profile_results.prof')
        print("\nProfile saved to: profile_results.prof")
        print("View with: python -m pstats profile_results.prof")
    else:
        # Serial mode without profiling
        run_batch(exam_config, writer, benchmark=enable_benchmark)
    
    # Save results
    if enable_benchmark:
        t0 = time.perf_counter()
    excel_writer.save()
    if enable_benchmark:
        excel_save_time = time.perf_counter() - t0
        print(f"\nExcel: load {excel_load_time*1000:.1f}ms, save {excel_save_time*1000:.1f}ms")
    
    print(f"\nDone: {exam_config.name}")


if __name__ == "__main__":
    # Required for Windows multiprocessing
    multiprocessing.freeze_support()
    main()
