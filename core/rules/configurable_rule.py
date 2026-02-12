"""
Configurable rule that composes multiple atomic check units.
"""

import time
from core.rules.base_rule import BaseRule
from core.checks.registry import get_check
from typing import List, Dict, Any

# COM error codes that are retriable (Word is alive but busy)
_RETRIABLE_COM_CODES = {
    -2147418111,  # RPC_E_CALL_REJECTED - Word busy, rejecting calls
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
}



class ConfigurableRule(BaseRule):
    """
    A rule composed of multiple atomic check units through configuration.
    
    This allows creating complex rules without writing new code, just by
    specifying which checks to run and how to score them.
    """
    
    def __init__(self, name: str, mark: int, checks: List[Dict[str, Any]], 
                 scoring: str = "all_or_nothing", min_pass: int = None, 
                 needs_review_on_error: bool = False,
                 always_review: bool = False,
                 review_on_partial: bool = False):
        """
        Initialize a configurable rule.
        
        Args:
            name: Display name for this rule
            mark: Points awarded (can be negative for deduction)
            checks: List of check configurations
                Example: [{"type": "margin", "params": {"side": "Top", "expected_cm": 3.0}}]
            scoring: Scoring mode:
                - "all_or_nothing": Full mark if all checks pass, 0 otherwise (default)
                - "at_least": Full mark if at least min_pass checks pass
                - "partial": Proportional mark based on pass rate (kept for flexibility)
            min_pass: Minimum number of checks that must pass (for "at_least" mode)
            needs_review_on_error: If True, any check error triggers manual review
            always_review: If True, always mark for manual review regardless of result
            review_on_partial: If True, mark for review when not full marks (skip review on full marks)
        """
        super().__init__(name=name, mark=mark)
        self.check_configs = checks
        self.scoring = scoring
        self.min_pass = min_pass
        self.needs_review_on_error = needs_review_on_error
        self.always_review = always_review
        self.review_on_partial = review_on_partial
        self.check_instances = []
        
        # Build check instances from configuration
        self._build_checks()
    
    def _build_checks(self):
        """Build check instances from configuration."""
        for check_config in self.check_configs:
            check_type = check_config.get("type")
            params = check_config.get("params", {})
            
            if not check_type:
                raise ValueError(f"Check configuration missing 'type' field: {check_config}")
            
            try:
                check_class = get_check(check_type)
                check_instance = check_class(**params)
                self.check_instances.append(check_instance)
            except Exception as e:
                self.logger.error(f"Failed to build check '{check_type}': {e}")
                raise ValueError(f"Invalid check configuration: {check_type} with params {params}") from e
    
    def run(self, doc) -> dict:
        """
        Execute all configured checks and compute score.
        
        Args:
            doc: Word document object
        
        Returns:
            dict: {
                "name": rule name,
                "mark": computed mark,
                "errors": list of error messages,
                "needs_review": whether manual review is needed
            }
        """
        # Handle placeholder rules (no checks configured)
        if not self.check_instances:
            return {
                "name": self.name,
                "mark": 0,
                "errors": [],
                "needs_review": False,
                "placeholder": True
            }
        
        errors = []
        passed_count = 0
        check_error_occurred = False
        
        # Run all checks
        for i, check in enumerate(self.check_instances):
            try:
                success, error_msg = self._run_check_with_retry(check, doc)
                
                if success:
                    passed_count += 1
                    self.logger.debug(f"Check {i+1}/{len(self.check_instances)} passed: {check}")
                else:
                    if error_msg:
                        errors.append(error_msg)
                    self.logger.debug(f"Check {i+1}/{len(self.check_instances)} failed: {check} - {error_msg}")
                
            except Exception as e:
                error_msg = f"Check execution error ({check}): {str(e)}"
                errors.append(error_msg)
                check_error_occurred = True
                self.logger.error(error_msg)
        
        # Calculate mark based on scoring mode
        mark = self._calculate_mark(passed_count, len(self.check_instances))
        
        # Determine if manual review is needed
        is_full_mark = (mark == self.mark)
        needs_review = (self.always_review
                        or (self.review_on_partial and not is_full_mark)
                        or (check_error_occurred and self.needs_review_on_error))
        
        return {
            "name": self.name,
            "mark": mark,
            "errors": errors,
            "needs_review": needs_review
        }
    
    def _run_check_with_retry(self, check, doc, max_retries=3):
        """
        Run a check with automatic retry for RPC_E_CALL_REJECTED errors.
        
        When Word is busy (e.g. under high concurrency), COM calls get rejected
        with -2147418111. This is transient - retrying after a brief wait works.
        
        Args:
            check: Check instance to execute
            doc: Word document COM object
            max_retries: Maximum retry attempts
            
        Returns:
            (success, error_msg) tuple from the check
        """
        for attempt in range(max_retries):
            try:
                return check.check(doc)
            except Exception as e:
                # Extract COM error code if available
                error_code = getattr(e, 'hresult', None)
                if error_code is None and hasattr(e, 'args') and len(e.args) >= 1:
                    if isinstance(e.args[0], int):
                        error_code = e.args[0]
                
                if error_code in _RETRIABLE_COM_CODES and attempt < max_retries - 1:
                    # Word is busy - wait with exponential backoff and retry
                    wait = 0.2 * (2 ** attempt)  # 0.2s, 0.4s, 0.8s
                    time.sleep(wait)
                    continue
                else:
                    raise  # Non-retriable or exhausted retries
        
        # Should not reach here, but just in case
        return check.check(doc)

    def _calculate_mark(self, passed: int, total: int) -> int:
        """
        Calculate mark based on scoring mode.
        
        Args:
            passed: Number of checks that passed
            total: Total number of checks
        
        Returns:
            Computed mark
        """
        if self.scoring == "all_or_nothing":
            # All checks must pass to get full mark
            return self.mark if passed == total else 0
        
        elif self.scoring == "at_least":
            # At least min_pass checks must pass to get full mark
            # Example: 4 checks, min_pass=3 → need at least 3 to pass
            if self.min_pass is None:
                self.logger.warning("'at_least' scoring requires min_pass parameter, defaulting to all_or_nothing")
                return self.mark if passed == total else 0
            
            return self.mark if passed >= self.min_pass else 0
        
        elif self.scoring == "partial":
            # Proportional scoring (kept for flexibility, but generally not used)
            if total == 0:
                return 0
            ratio = passed / total
            return int(self.mark * ratio)
        
        else:
            self.logger.warning(f"Unknown scoring mode '{self.scoring}', defaulting to all_or_nothing")
            return self.mark if passed == total else 0
