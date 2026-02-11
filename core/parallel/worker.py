# core/parallel/worker.py
"""
Word worker process management for pooled Word instances.
方案 A: Word 实例复用 + 方案 B: 信号量控制
"""

import time
import atexit

# Worker process global state
_worker_word_app = None
_worker_semaphore = None
_worker_initialized = False


def init_word_worker(semaphore, init_barrier):
    """
    Initialize Word instance when worker process starts.
    
    Uses a Barrier to ensure ALL workers finish initializing before
    ANY tasks are dispatched. This prevents new Word instance creation
    from interfering with active COM processing.
    
    Args:
        semaphore: Multiprocessing semaphore to limit concurrent Word instances
        init_barrier: Multiprocessing barrier - all workers wait until everyone is ready
    """
    global _worker_word_app, _worker_semaphore, _worker_initialized
    
    # Disable file logging in worker process to avoid file lock conflicts
    try:
        from core.logging.logger import disable_file_logging
        disable_file_logging()
    except:
        pass
    
    _worker_semaphore = semaphore
    
    try:
        # Initialize COM in this process
        import pythoncom
        pythoncom.CoInitialize()
        
        # Acquire semaphore before creating Word instance
        _worker_semaphore.acquire()
        
        # Create Word application instance
        import win32com.client
        _worker_word_app = win32com.client.Dispatch("Word.Application")
        _worker_word_app.Visible = False
        
        # Brief stabilization wait
        time.sleep(0.3)
        
        # Close any existing documents
        try:
            while _worker_word_app.Documents.Count > 0:
                _worker_word_app.Documents(1).Close(SaveChanges=False)
        except:
            pass
        
        _worker_initialized = True
        
        # Register cleanup function
        atexit.register(cleanup_word_worker)
        
    except Exception as e:
        # If initialization failed, release semaphore
        if _worker_semaphore:
            _worker_semaphore.release()
        _worker_initialized = False
        # Don't raise - let barrier proceed so other workers aren't stuck
    
    # Wait for ALL workers to finish initialization before accepting tasks.
    # This prevents new Word instance creation from destabilizing
    # existing COM connections during active document processing.
    try:
        init_barrier.wait(timeout=60)
    except:
        pass  # Continue even if barrier times out or breaks


def cleanup_word_worker():
    """
    Clean up Word instance when worker process exits.
    
    This function:
    1. Closes all open documents
    2. Quits Word application
    3. Releases the semaphore
    4. Uninitializes COM
    """
    global _worker_word_app, _worker_semaphore, _worker_initialized
    
    # Close Word application
    if _worker_word_app:
        try:
            # Close all documents
            while _worker_word_app.Documents.Count > 0:
                try:
                    _worker_word_app.Documents(1).Close(SaveChanges=False)
                except:
                    break
            
            # Quit Word
            _worker_word_app.Quit()
        except:
            pass  # Ignore cleanup errors
        finally:
            _worker_word_app = None
    
    # Release semaphore
    if _worker_semaphore and _worker_initialized:
        try:
            _worker_semaphore.release()
        except:
            pass
    
    # Uninitialize COM
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except:
        pass
    
    _worker_initialized = False


def get_worker_state():
    """Get current worker state for debugging."""
    return {
        'initialized': _worker_initialized,
        'has_word_app': _worker_word_app is not None,
        'has_semaphore': _worker_semaphore is not None,
    }


def recover_word_instance():
    """Recreate the Word instance if it has died."""
    global _worker_word_app, _worker_initialized
    
    # Close old instance if any
    if _worker_word_app:
        try:
            _worker_word_app.Quit()
        except:
            pass
        _worker_word_app = None
    
    try:
        import win32com.client
        _worker_word_app = win32com.client.Dispatch("Word.Application")
        _worker_word_app.Visible = False
        time.sleep(0.3)
        _worker_initialized = True
        return True
    except Exception:
        _worker_initialized = False
        return False


def get_worker_word_app():
    """Get the pooled Word application instance. Recovers if dead."""
    global _worker_word_app, _worker_initialized
    
    if not _worker_initialized or _worker_word_app is None:
        raise RuntimeError("Word worker not initialized")
    
    # Quick health check - recover if dead
    try:
        _ = _worker_word_app.Version
    except:
        if not recover_word_instance():
            raise RuntimeError("Word worker died and recovery failed")
    
    return _worker_word_app
