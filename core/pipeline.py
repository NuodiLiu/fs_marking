# core/pipeline.py

import os
import traceback
from core.document_loader import load_word_document
from core.rule_engine import evaluate
from core.utils.utils import validate_zid
from core.utils.result_utils import merge_results
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)

SUBMISSION_DIR = os.path.join(ROOT_DIR, "data", "submissions")
LOG_DIR = os.path.join(ROOT_DIR, "logs")

def find_word_file(student_folder):
    word_files = []
    for root, dirs, files in os.walk(student_folder):
        for file in files:
            if file.startswith(".") or file.startswith("~"):
                continue
            if file.lower().endswith((".doc", ".docx")):
                word_files.append(os.path.join(root, file))

    if len(word_files) == 0:
        raise FileNotFoundError("❌ No Word document found.")
    elif len(word_files) > 1:
        raise ValueError("❌ Multiple Word files found.")
    return word_files[0]
    

def run_batch(exam_config, writer=None, benchmark=False):
    """
    Run batch evaluation for all student submissions.
    
    Args:
        exam_config: ExamConfig object with rules and output configuration
        writer: Optional ResultWriter for outputting results
        benchmark: If True, collect detailed timing statistics
    """
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)

    total_start_time = time.time()
    timings = []
    
    # Benchmark statistics
    if benchmark:
        bench_stats = {
            'find_file': [],
            'load_word': [],
            'evaluate': [],
            'merge': [],
            'close_word': [],
            'write_results': []
        }

    for zid_folder in os.listdir(SUBMISSION_DIR):
        student_path = os.path.join(SUBMISSION_DIR, zid_folder)
        if not os.path.isdir(student_path):
            continue

        if not validate_zid(zid_folder):
            print(f"⚠️ Skipped invalid folder name: {zid_folder}")
            continue

        print(f"🔍 Checking {zid_folder}...")
        start_time = time.time()

        try:
            # 1. Find word file
            if benchmark:
                t0 = time.perf_counter()
            word_path = find_word_file(student_path)
            if benchmark:
                bench_stats['find_file'].append(time.perf_counter() - t0)
            
            # 2. Load Word document
            if benchmark:
                t0 = time.perf_counter()
            doc, word_app = load_word_document(word_path)
            if benchmark:
                bench_stats['load_word'].append(time.perf_counter() - t0)

            # 3. Evaluate rules
            if benchmark:
                t0 = time.perf_counter()
            result = evaluate(doc, exam_config.rules)
            if benchmark:
                bench_stats['evaluate'].append(time.perf_counter() - t0)
            
            # 4. Apply merge rules if configured
            if exam_config.merge_rules:
                if benchmark:
                    t0 = time.perf_counter()
                result["results"] = merge_results(result["results"], exam_config.merge_rules)
                if benchmark:
                    bench_stats['merge'].append(time.perf_counter() - t0)

            # 5. Close Word
            if benchmark:
                t0 = time.perf_counter()
            doc.Close(False)
            word_app.Quit()
            del doc
            del word_app
            if benchmark:
                bench_stats['close_word'].append(time.perf_counter() - t0)

            # 6. Write results
            if writer is not None:
                if benchmark:
                    t0 = time.perf_counter()
                writer.write(zid_folder, result)
                if benchmark:
                    bench_stats['write_results'].append(time.perf_counter() - t0)

            elapsed = time.time() - start_time
            timings.append(elapsed)
            # print(f"✅ Finished {zid_folder}: {result['total']} marks ({elapsed:.2f}s)\n\n")

        except Exception as e:
            elapsed = time.time() - start_time
            timings.append(elapsed)
            error_msg = f"❌ Failed {zid_folder} ({elapsed:.2f}s): {e}"
            print(error_msg)
            with open(os.path.join(LOG_DIR, f"{zid_folder}_error.log"), "w", encoding="utf-8") as f:
                f.write(error_msg + "\n")
                f.write(traceback.format_exc())

    total_time = time.time() - total_start_time
    count = len(timings)
    avg_time = sum(timings) / count if count > 0 else 0
    print(f"\n⏱️ Finished batch in {total_time:.2f}s")
    print(f"📊 Average time per document: {avg_time:.2f}s")
    
    # Print benchmark statistics
    if benchmark and count > 0:
        print("\n" + "="*80)
        print("⏱️  BENCHMARK RESULTS (per document average)")
        print("="*80)
        
        def avg_ms(times):
            return (sum(times) / len(times) * 1000) if times else 0
        
        def total_s(times):
            return sum(times) if times else 0
        
        def pct(times, total):
            return (sum(times) / total * 100) if total > 0 else 0
        
        total_bench = sum(sum(v) for v in bench_stats.values())
        
        print(f"{'Operation':<25} {'Avg (ms)':<12} {'Total (s)':<12} {'% of Total':<12}")
        print("-" * 80)
        print(f"{'1. Find Word File':<25} {avg_ms(bench_stats['find_file']):<12.2f} {total_s(bench_stats['find_file']):<12.2f} {pct(bench_stats['find_file'], total_bench):<12.1f}")
        print(f"{'2. Load Word Document':<25} {avg_ms(bench_stats['load_word']):<12.2f} {total_s(bench_stats['load_word']):<12.2f} {pct(bench_stats['load_word'], total_bench):<12.1f}")
        print(f"{'3. Evaluate Rules':<25} {avg_ms(bench_stats['evaluate']):<12.2f} {total_s(bench_stats['evaluate']):<12.2f} {pct(bench_stats['evaluate'], total_bench):<12.1f}")
        if bench_stats['merge']:
            print(f"{'4. Merge Results':<25} {avg_ms(bench_stats['merge']):<12.2f} {total_s(bench_stats['merge']):<12.2f} {pct(bench_stats['merge'], total_bench):<12.1f}")
        print(f"{'5. Close Word':<25} {avg_ms(bench_stats['close_word']):<12.2f} {total_s(bench_stats['close_word']):<12.2f} {pct(bench_stats['close_word'], total_bench):<12.1f}")
        print(f"{'6. Write Results (FS)':<25} {avg_ms(bench_stats['write_results']):<12.2f} {total_s(bench_stats['write_results']):<12.2f} {pct(bench_stats['write_results'], total_bench):<12.1f}")
        print("-" * 80)
        print(f"{'TOTAL (measured)':<25} {total_bench/count*1000:<12.2f} {total_bench:<12.2f} {'100.0':<12}")
        print("="*80)