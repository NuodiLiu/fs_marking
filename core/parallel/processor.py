# core/parallel/processor.py
"""
Document processing functions for parallel execution.
"""

import time
import random
from typing import Tuple, Dict, Any, Optional

from core.document_loader import load_word_document
from core.rule_engine import evaluate
from core.utils.result_utils import merge_results
from .worker import get_worker_word_app, recover_word_instance


def process_document_with_pooled_word(args: Tuple) -> Tuple[str, Optional[Dict], float, bool, str]:
    """
    Process a single document using the pooled Word instance.
    
    This is the optimized processing function that reuses the Word
    instance created during worker initialization.
    
    Args:
        args: (zid, word_path, config_module_name, merge_rules)
    
    Returns:
        (zid, result_dict, elapsed_time, success, error_message)
    """
    zid, word_path, config_module_name, merge_rules = args
    start_time = time.time()
    doc = None
    
    try:
        # Get the pooled Word application (auto-recovers if dead)
        word_app = get_worker_word_app()
        
        # Load configuration
        from core.config_loader import ConfigLoader
        exam_config = ConfigLoader.load_from_module(config_module_name)
        
        # Open document using existing Word instance with retry + recovery
        max_retries = 2
        last_error = None
        
        for attempt in range(max_retries):
            try:
                doc = word_app.Documents.Open(word_path)
                break
            except Exception as e:
                last_error = e
                if attempt < max_retries - 1:
                    time.sleep(0.3)
                    # Try cleanup, then recover Word instance
                    try:
                        while word_app.Documents.Count > 0:
                            word_app.Documents(1).Close(SaveChanges=False)
                    except:
                        pass
                    # Word may be dead - try to recover
                    if recover_word_instance():
                        word_app = get_worker_word_app()
                else:
                    raise RuntimeError(f"Cannot open document: {e}")
        
        # Evaluate rules
        result = evaluate(doc, exam_config.rules)
        
        # Apply merge rules if configured
        if merge_rules:
            result["results"] = merge_results(result["results"], merge_rules)
        
        # Close document (but keep Word instance)
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass
        
        elapsed = time.time() - start_time
        return (zid, result, elapsed, True, "")
        
    except Exception as e:
        # Try to close document on error
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass
        
        elapsed = time.time() - start_time
        error_msg = f"{type(e).__name__}: {str(e)}"
        return (zid, None, elapsed, False, error_msg)


def process_single_document(args: Tuple) -> Tuple[str, Optional[Dict], float, bool, str]:
    """
    Worker function for parallel processing (legacy, creates new Word instance).
    
    This function creates a new Word instance for each document.
    Only used as fallback - prefer process_document_with_pooled_word.
    
    Args:
        args: (zid, word_path, config_module_name, merge_rules)
    
    Returns:
        (zid, result_dict, elapsed_time, success, error_message)
    """
    zid, word_path, config_module_name, merge_rules = args
    
    start_time = time.time()
    doc = None
    word_app = None
    
    try:
        # Initialize COM for this process (required for Word automation)
        import pythoncom
        pythoncom.CoInitialize()
        
        try:
            # Add random delay to stagger Word instance creation
            time.sleep(random.uniform(0.1, 0.5))
            
            # Load configuration in worker process
            from core.config_loader import ConfigLoader
            exam_config = ConfigLoader.load_from_module(config_module_name)
            
            # Retry logic for Word COM initialization
            max_retries = 3
            last_error = None
            
            for attempt in range(max_retries):
                try:
                    # Load and process document
                    doc, word_app = load_word_document(word_path)
                    break
                except Exception as e:
                    last_error = e
                    if attempt < max_retries - 1:
                        # Wait before retry with exponential backoff
                        time.sleep(0.5 * (2 ** attempt))
                        # Try to clean up any partial COM objects
                        try:
                            import gc
                            gc.collect()
                        except:
                            pass
                    else:
                        raise last_error
            
            # Evaluate rules
            result = evaluate(doc, exam_config.rules)
            
            # Apply merge rules if configured
            if merge_rules:
                result["results"] = merge_results(result["results"], merge_rules)
            
            # Close Word safely
            try:
                if doc:
                    doc.Close(False)
                if word_app:
                    word_app.Quit()
            except:
                pass  # Ignore cleanup errors
            
            del doc
            del word_app
            
            elapsed = time.time() - start_time
            return (zid, result, elapsed, True, "")
            
        finally:
            # Ensure Word is cleaned up
            try:
                if doc:
                    doc.Close(False)
            except:
                pass
            try:
                if word_app:
                    word_app.Quit()
            except:
                pass
            
            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            
    except Exception as e:
        elapsed = time.time() - start_time
        error_msg = f"{type(e).__name__}: {str(e)}"
        return (zid, None, elapsed, False, error_msg)
