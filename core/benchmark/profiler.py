# core/benchmark/profiler.py
"""
Benchmark profiling for pipeline operations.
"""

import time
from contextlib import contextmanager
from typing import Dict, List


class BenchmarkProfiler:
    """
    Context manager for collecting detailed performance statistics.
    
    Usage:
        profiler = BenchmarkProfiler()
        
        with profiler.measure('operation_name'):
            # ... do work ...
        
        profiler.print_report()
    """
    
    def __init__(self):
        self.stats: Dict[str, List[float]] = {}
        self.enabled = True
    
    @contextmanager
    def measure(self, operation: str):
        """
        Context manager to measure operation time.
        
        Args:
            operation: Name of the operation being measured
        """
        if not self.enabled:
            yield
            return
        
        start = time.perf_counter()
        try:
            yield
        finally:
            elapsed = time.perf_counter() - start
            if operation not in self.stats:
                self.stats[operation] = []
            self.stats[operation].append(elapsed)
    
    def record(self, operation: str, duration: float):
        """
        Manually record a measurement.
        
        Args:
            operation: Name of the operation
            duration: Time in seconds
        """
        if not self.enabled:
            return
        
        if operation not in self.stats:
            self.stats[operation] = []
        self.stats[operation].append(duration)
    
    def get_stats(self) -> Dict[str, Dict[str, float]]:
        """
        Get computed statistics for all operations.
        
        Returns:
            Dict mapping operation name to stats (avg, total, count, pct)
        """
        if not self.stats:
            return {}
        
        total_time = sum(sum(times) for times in self.stats.values())
        
        result = {}
        for operation, times in self.stats.items():
            count = len(times)
            total = sum(times)
            avg = total / count if count > 0 else 0
            pct = (total / total_time * 100) if total_time > 0 else 0
            
            result[operation] = {
                'avg': avg,
                'total': total,
                'count': count,
                'pct': pct,
            }
        
        return result
    
    def print_report(self, title="BENCHMARK RESULTS"):
        """
        Print a formatted benchmark report.
        
        Args:
            title: Report title
        """
        if not self.stats:
            print("No benchmark data collected.")
            return
        
        stats = self.get_stats()
        
        print("\n" + "=" * 80)
        print(f"{title} (per document average)")
        print("=" * 80)
        print(f"{'Operation':<30} {'Avg (ms)':<12} {'Total (s)':<12} {'% of Total':<12}")
        print("-" * 80)
        
        for operation, data in stats.items():
            avg_ms = data['avg'] * 1000
            total_s = data['total']
            pct = data['pct']
            print(f"{operation:<30} {avg_ms:<12.2f} {total_s:<12.2f} {pct:<12.1f}")
        
        print("-" * 80)
        
        # Calculate totals
        total_time = sum(data['total'] for data in stats.values())
        count = max(data['count'] for data in stats.values())
        avg_total_ms = (total_time / count * 1000) if count > 0 else 0
        
        print(f"{'TOTAL (measured)':<30} {avg_total_ms:<12.2f} {total_time:<12.2f} {'100.0':<12}")
        print("=" * 80)
    
    def reset(self):
        """Clear all collected statistics."""
        self.stats.clear()
    
    def disable(self):
        """Disable profiling."""
        self.enabled = False
    
    def enable(self):
        """Enable profiling."""
        self.enabled = True


def create_operation_names():
    """Standard operation names for pipeline benchmarking."""
    return {
        'FIND_FILE': 'Find Word File',
        'LOAD_WORD': 'Load Word Document',
        'EVALUATE': 'Evaluate Rules',
        'MERGE': 'Merge Results',
        'CLOSE_WORD': 'Close Word',
        'WRITE_RESULTS': 'Write Results (FS)',
    }
