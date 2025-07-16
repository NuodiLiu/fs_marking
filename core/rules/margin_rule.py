from core.rules.base_rule import BaseRule

class MarginRule(BaseRule):
    def __init__(self, mark=1):
        super().__init__(name="Margin Check", mark=mark)

    def run(self, doc):
        try:
            section = doc.Sections(1).PageSetup
            # Word uses points (pt), 1 cm ≈ 28.35 pt
            expected_margin_pt = 2.5 * 28.35
            tolerance = 0.01 # +- 0.00035cm to avoid converting error

            margins = {
                "Top": section.TopMargin,
                "Bottom": section.BottomMargin,
                "Left": section.LeftMargin,
                "Right": section.RightMargin
            }

            errors = []
            for side, actual in margins.items():
                if abs(actual - expected_margin_pt) > tolerance:
                    actual_cm = actual / 28.35
                    errors.append(
                        f"{side} margin is {actual_cm:.2f}cm, expected 2.5cm"
                    )

            if errors:
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": errors,
                    "needs_review": False
                }

            return {
                "name": self.name,
                "mark": self.mark,
                "errors": [],
                "needs_review": False
            }

        except Exception as e:
            return {
                "name": self.name,
                "mark": 0,
                "errors": [f"Margin check error: {str(e)}"],
                "needs_review": True
            }
