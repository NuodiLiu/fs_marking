# core/parallel/executor.py
"""
Parallel execution engine for batch document processing.
"""

import os
import multiprocessing
from typing import Optional

from core.utils.utils import validate_zid
from .worker import init_word_worker
from .processor import process_document_with_pooled_word

# Get default directories
def _get_default_dirs():
    """Get default submission and log directories."""
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    root_dir = os.path.dirname(base_dir)
    return (
        os.path.join(root_dir, "data", "submissions"),
        os.path.join(root_dir, "logs")
    )


def _has_com_errors(result_dict):
    """
    Check if evaluation results contain COM errors.
    
    Args:
        result_dict: Result dictionary from evaluate()
        
    Returns:
        True if COM errors detected, False otherwise
    """
    if result_dict is None:
        return False
    
    results = result_dict.get("results", [])
    for rule_result in results:
        errors = rule_result.get("errors", [])
        for error in errors:
            error_msg = str(error).lower()
            # Check for common COM error indicators
            if any(indicator in error_msg for indicator in [
                "被调用的对象已与其客户端断开连接",
                "被呼叫方拒绝接收呼叫",
                "rpc",
                "com_error",
                "-2147417848",  # RPC_E_DISCONNECTED
                "-2147418111",  # RPC_E_CALL_REJECTED
            ]):
                return True
    return False


def _is_retriable_error(error_msg):
    """
    Determine if an error is retriable (worth trying again in serial mode).
    
    Args:
        error_msg: Error message string
        
    Returns:
        True if error is retriable, False otherwise
    """
    error_lower = error_msg.lower()
    
    # Retriable errors (COM/RPC issues, temporary failures)
    retriable_indicators = [
        "com_error",
        "rpc",
        "被调用的对象已与其客户端断开连接",
        "被呼叫方拒绝接收呼叫",
        "所要求的对象不存在",
        "rpc 服务器不可用",
        "服务器运行失败",
        "-2147417848",  # RPC_E_DISCONNECTED
        "-2147418111",  # RPC_E_CALL_REJECTED
        "-2147417836",  # Object doesn't exist
        "-2147023174",  # RPC server unavailable
        "-2146959355",  # Server execution failed
        "cannot open document",
        "timeout",
    ]
    
    # Non-retriable errors (file not found, multiple files, etc.)
    non_retriable_indicators = [
        "no word document found",
        "multiple word files found",
        "file does not exist",
        "filenotfounderror",
    ]
    
    # Check if non-retriable
    if any(indicator in error_lower for indicator in non_retriable_indicators):
        return False
    
    # Check if retriable
    if any(indicator in error_lower for indicator in retriable_indicators):
        return True
    
    # Default: don't retry unknown errors (conservative)
    return False


def run_batch_parallel(exam_config, writer=None, benchmark=False, num_workers=None, 
                      submission_dir=None, log_dir=None):
    """
    Run batch evaluation using multiprocessing with pooled Word instances.
    
    Implementation:
    - Method A: Worker process reuses a single Word instance for all documents
    - Method B: Semaphore controls maximum concurrent Word instances
    
    Args:
        exam_config: ExamConfig object with rules and output configuration
        writer: Optional ResultWriter for outputting results
        benchmark: If True, collect detailed timing statistics
        num_workers: Number of worker processes (default: 2, max recommended: 4)
        submission_dir: Directory containing student submissions (default: data/submissions)
        log_dir: Directory for error logs (default: logs)
    
    Returns:
        List of timing measurements for each document
    """
    import time
    
    # Set default directories if not provided
    if submission_dir is None or log_dir is None:
        default_submission, default_log = _get_default_dirs()
        if submission_dir is None:
            submission_dir = default_submission
        if log_dir is None:
            log_dir = default_log
    
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    if num_workers is None:
        # Default to 2 workers (conservative for Word COM)
        # You can increase with -j flag: python main.py exam_feb11 --parallel -j 8
        num_workers = 2
    
    # Show configuration info
    if num_workers > 4:
        print(f"Using {num_workers} workers (high concurrency mode)")
        print(f"  Note: Word COM typically works best with 2-4 workers due to RPC/filesystem limits")
        print(f"  If you see errors, try reducing with: -j 4\n")
    
    # Create semaphore to control maximum concurrent Word instances
    max_concurrent_word = num_workers
    word_semaphore = multiprocessing.Semaphore(max_concurrent_word)
    
    # Barrier ensures ALL workers finish init before ANY tasks are dispatched.
    # This prevents new Word instance creation from destabilizing
    # existing COM/RPC connections during active document processing.
    init_barrier = multiprocessing.Barrier(num_workers)
    
    total_start_time = time.time()
    
    # Prepare task list
    tasks = []
    for zid_folder in os.listdir(submission_dir):
        student_path = os.path.join(submission_dir, zid_folder)
        if not os.path.isdir(student_path):
            continue
        
        if not validate_zid(zid_folder):
            print(f"Skipped invalid folder name: {zid_folder}")
            continue
        
        try:
            # Import here to avoid circular dependency
            from ..pipeline import find_word_file
            word_path = find_word_file(student_path)
            # Pass config module name instead of object (for pickling)
            tasks.append((zid_folder, word_path, exam_config.source_module, exam_config.merge_rules))
        except Exception as e:
            print(f"Skipped {zid_folder}: {e}")
    
    if not tasks:
        print("No valid documents to process")
        return []
    
    print(f"Processing {len(tasks)} documents with {num_workers} workers...")
    print(f"Starting {num_workers} Word instances (this may take a few seconds)...\n")
    
    # Process in parallel with pooled Word instances
    results = []
    timings = []
    errors = []
    retry_queue = []  # All tasks that need retry (COM errors + retriable failures)
    
    try:
        # Use initializer to create Word instance when each worker starts
        # Lock ensures workers initialize Word instances one at a time
        with multiprocessing.Pool(
            processes=num_workers,
            initializer=init_word_worker,
            initargs=(word_semaphore, init_barrier)
        ) as pool:
            # Use imap_unordered for progress tracking
            for result in pool.imap_unordered(process_document_with_pooled_word, tasks):
                zid, result_dict, elapsed, success, error_msg = result
                timings.append(elapsed)
                
                if success:
                    # Check if result contains COM errors
                    if _has_com_errors(result_dict):
                        print(f"  [RETRY] {zid} ({elapsed:.2f}s) - has COM errors, will retry")
                        # Find original task for retry
                        for task in tasks:
                            if task[0] == zid:
                                retry_queue.append(task)
                                break
                    else:
                        print(f"  [OK] {zid} ({elapsed:.2f}s)")
                        results.append((zid, result_dict))
                else:
                    # Task completely failed
                    if _is_retriable_error(error_msg):
                        print(f"  [RETRY] {zid} ({elapsed:.2f}s) - retriable error, will retry")
                        # Find original task for retry
                        for task in tasks:
                            if task[0] == zid:
                                retry_queue.append(task)
                                break
                    else:
                        print(f"  [FAIL] {zid} ({elapsed:.2f}s): {error_msg}")
                        errors.append((zid, error_msg, elapsed))
        
        # Pool context manager will automatically call cleanup_word_worker
        
    except KeyboardInterrupt:
        print("\nInterrupted by user. Cleaning up...")
        raise
    except Exception as e:
        print(f"\nParallel processing failed: {e}")
        print(f"Tip: Try reducing workers with -j 1 or use serial mode without --parallel")
        raise
    
    # Retry failed tasks with single worker (serial mode)
    if retry_queue:
        print(f"\n{'='*60}")
        print(f"RETRY PHASE: Processing {len(retry_queue)} documents in serial mode")
        print(f"{'='*60}\n")
        
        retry_semaphore = multiprocessing.Semaphore(1)
        retry_barrier = multiprocessing.Barrier(1)
        
        with multiprocessing.Pool(
            processes=1,  # Single worker for stability
            initializer=init_word_worker,
            initargs=(retry_semaphore, retry_barrier)
        ) as retry_pool:
            for result in retry_pool.imap(process_document_with_pooled_word, retry_queue):
                zid, result_dict, elapsed, success, error_msg = result
                timings.append(elapsed)
                
                if success and not _has_com_errors(result_dict):
                    print(f"  [OK] {zid} ({elapsed:.2f}s) - retry successful")
                    results.append((zid, result_dict))
                else:
                    print(f"  [FAIL] {zid} ({elapsed:.2f}s) - retry failed: {error_msg if not success else 'COM errors persist'}")
                    errors.append((zid, error_msg if not success else "COM errors after retry", elapsed))
    
    # Write results (single-threaded, no conflicts)
    for zid, result_dict in results:
        if writer is not None:
            writer.write(zid, result_dict)
    
    # Write error logs
    if log_dir:
        for zid, error_msg, elapsed in errors:
            error_log_path = os.path.join(log_dir, f"{zid}_error.log")
            with open(error_log_path, "w", encoding="utf-8") as f:
                f.write(f"Failed {zid} ({elapsed:.2f}s)\n")
                f.write(f"{error_msg}\n")
    
    total_time = time.time() - total_start_time
    successful_count = len(results)
    total_count = successful_count + len(errors)
    avg_time = total_time / total_count if total_count > 0 else 0
    
    print(f"\n{'='*60}")
    print(f"BATCH PROCESSING COMPLETE")
    print(f"{'='*60}")
    print(f"Total wall time : {total_time:.2f}s")
    print(f"Avg per document: {avg_time:.2f}s")
    print(f"Successful      : {successful_count}/{total_count}")
    print(f"Failed          : {len(errors)}")
    
    if retry_queue:
        retry_success = len([zid for zid, _ in results if any(zid == task[0] for task in retry_queue)])
        print(f"Retried         : {len(retry_queue)} | Recovered: {retry_success}")
    
    print(f"{'='*60}")
    
    return timings
