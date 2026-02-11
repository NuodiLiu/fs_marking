# core/parallel/__init__.py
"""
Parallel processing module for Word document batch evaluation.
"""

from .worker import init_word_worker, cleanup_word_worker
from .processor import process_document_with_pooled_word, process_single_document
from .executor import run_batch_parallel

__all__ = [
    'init_word_worker',
    'cleanup_word_worker',
    'process_document_with_pooled_word',
    'process_single_document',
    'run_batch_parallel',
]
