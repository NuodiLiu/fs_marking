# core/rules/footnote_rules/footnote_on_habitat_rule.py

from core.rules.base_rule import BaseRule
from core.utils import fuzzy_match

class FootnoteOnHabitatRule(BaseRule):
    def __init__(self, mark=1, keyword="habitat", expected_note="Natural home or environment"):
        super().__init__("Footnote on Habitat Rule", mark)
        self.keyword = keyword
        self.expected_note = expected_note

    def run(self, doc):
        try:
            errors = []
            found = False

            for word_range in doc.Words:
                word_text = word_range.Text.strip().rstrip(".,;:!?")
                if word_text.lower() == self.keyword.lower():
                    # 找到首次 habitat，检查是否有脚注
                    if word_range.Footnotes.Count == 0:
                        errors.append("'habitat' does not have a footnote.")
                    else:
                        # 检查内容模糊匹配
                        note = word_range.Footnotes(1).Range.Text.strip()
                        if not fuzzy_match(self.expected_note, note, threshold=0.75):
                            errors.append(f"Footnote content may be incorrect: '{note}'")

                    found = True
                    break

            if not found:
                errors.append("Could not find the word 'habitat' in the document.")

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
                "errors": [f"Error checking habitat footnote: {str(e)}"],
                "needs_review": True
            }
