# core/rules/combined_style_rule.py

from core.rules.base_rule import BaseRule
from core.rules.style_setting_rules.style_checkpoints_list import get_all_style_checkpoints

class CombinedStyleRule(BaseRule):
    def __init__(self):
        super().__init__("Combined Style Rule", 2)

    def run(self, doc):
        checkpoints = get_all_style_checkpoints()
        total_passed = 0
        errors = []
        exception_occurred = False

        for checkpoint in checkpoints:
            try:
                style_obj = doc.Styles(checkpoint.style_name)
                if checkpoint.run(style_obj):
                    total_passed += 1
                else:
                    errors.append(f"[{checkpoint.style_name}] {checkpoint.description}")
            except Exception as e:
                errors.append(f"[{checkpoint.style_name}] Error running check: {checkpoint.description} ({e})")
                exception_occurred = True

        mark = 2 if total_passed == len(checkpoints) else (1 if total_passed >= 6 else 0)

        return {
            "name": self.name,
            "mark": mark if not exception_occurred else 0,
            "errors": errors,
            "needs_review": exception_occurred
        }
