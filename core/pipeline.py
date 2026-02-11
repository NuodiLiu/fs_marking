# core/pipeline.py
"""
Main pipeline for batch document processing.

This module provides the main entry points for processing Word documents:
- find_word_file: Locate Word document in submission folder
- run_batch: Serial processing (with optional benchmarking)
- run_batch_parallel: Parallel processing (imported from parallel module)

For implementation details:
- See core.parallel for parallel processing
- See core.benchmark for profiling utilities
"""

import os
import traceback
import time
from typing import Optional

from core.document_loader import load_word_document
from core.rule_engine import evaluate
from core.utils.utils import validate_zid
from core.utils.result_utils import merge_results
from core.benchmark import BenchmarkProfiler

# Directory constants
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
SUBMISSION_DIR = os.path.join(ROOT_DIR, "data", "submissions")
LOG_DIR = os.path.join(ROOT_DIR, "logs")

# Import parallel processing function
from core.parallel import run_batch_parallel  # noqa: E402


def find_word_file(student_folder):
    """
    Find the Word document in a student submission folder.
    
    Args:
        student_folder: Path to student submission folder
        
    Returns:
        Path to the Word document
        
    Raises:
        FileNotFoundError: No Word document found
        ValueError: Multiple Word documents found
    """
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
    Run batch evaluation for all student submissions (serial processing).
    
    Args:
        exam_config: ExamConfig object with rules and output configuration
        writer: Optional ResultWriter for outputting results
        benchmark: If True, collect detailed timing statistics
        
    Returns:
        List of timing measurements for each document
    """
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)

    total_start_time = time.time()
    timings = []
    
    # Initialize benchmark profiler
    profiler = BenchmarkProfiler() if benchmark else None

    for zid_folder in os.listdir(SUBMISSION_DIR):
        student_path = os.path.join(SUBMISSION_DIR, zid_folder)
        if not os.path.isdir(student_path):
            continue

        if not validate_zid(zid_folder):
            continue

        start_time = time.time()

        try:
            # 1. Find word file
            with profiler.measure('Find Word File') if profiler else _null_context():
                word_path = find_word_file(student_path)
            
            # 2. Load Word document
            with profiler.measure('Load Word Document') if profiler else _null_context():
                doc, word_app = load_word_document(word_path)

            # 3. Evaluate rules
            with profiler.measure('Evaluate Rules') if profiler else _null_context():
                result = evaluate(doc, exam_config.rules)
            
            # 4. Apply merge rules if configured
            if exam_config.merge_rules:
                with profiler.measure('Merge Results') if profiler else _null_context():
                    result["results"] = merge_results(result["results"], exam_config.merge_rules)

            # 5. Close Word
            with profiler.measure('Close Word') if profiler else _null_context():
                doc.Close(False)
                word_app.Quit()
                del doc
                del word_app

            # 6. Write results
            if writer is not None:
                with profiler.measure('Write Results (FS)') if profiler else _null_context():
                    writer.write(zid_folder, result)

            elapsed = time.time() - start_time
            timings.append(elapsed)

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
    print(f"\nFinished batch in {total_time:.2f}s (avg {avg_time:.2f}s/doc)")
    
    # Print benchmark statistics
    if profiler:
        profiler.print_report()
    
    return timings


def _null_context():
    """Null context manager for when benchmarking is disabled."""
    from contextlib import nullcontext
    return nullcontext()
