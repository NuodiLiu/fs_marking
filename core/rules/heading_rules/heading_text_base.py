# core/rules/strict_heading_base.py
from core.rules.base_rule import BaseRule
from core.utils.utils import fuzzy_match

class StrictHeadingRuleBase(BaseRule):
    def __init__(self, name, mark, expected_headings, style_name):
        super().__init__(name, mark)
        self.expected_headings = expected_headings
        self.style_name = style_name.lower()

    def run(self, doc):
        try:
            heading_paragraphs = []
            unmatched_headings = self.expected_headings.copy()
            unexpected_headings = []

            for para in doc.Paragraphs:
                text = para.Range.Text.strip()
                style = para.Range.Style

                if not text:
                    continue

                style_actual = style if isinstance(style, str) else style.NameLocal
                if style_actual.lower() == self.style_name:
                    heading_paragraphs.append(text)

            matched = []

            for actual_heading in heading_paragraphs:
                match_found = False
                for expected_heading in unmatched_headings:
                    if fuzzy_match(expected_heading, actual_heading, threshold=0.85):
                        matched.append(expected_heading)
                        unmatched_headings.remove(expected_heading)
                        match_found = True
                        break
                if not match_found:
                    unexpected_headings.append(actual_heading)

            errors = []
            if unmatched_headings:
                errors.append(f"Missing expected {self.style_name.title()}: {unmatched_headings}")
            if unexpected_headings:
                errors.append(f"Found unexpected {self.style_name.title()}: {unexpected_headings}")

            return {
                "name": self.name,
                "mark": 0 if errors else self.mark,
                "errors": errors,
                "needs_review": False
            }

        except Exception as e:
            return {
                "name": self.name,
                "mark": 0,
                "errors": [f"Error during {self.style_name} check: {str(e)}"],
                "needs_review": True
            }
