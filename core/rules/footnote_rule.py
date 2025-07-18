# core/rules/footnote_rules/footnote_on_habitat_rule.py

from core.rules.base_rule import BaseRule
from core.utils.utils import fuzzy_match

class FootnoteOnHabitatRule(BaseRule):
    def __init__(self, mark=1, keyword="habitat", expected_note="Natural home or environment"):
        super().__init__("Footnote on Habitat Rule", mark)
        self.keyword = keyword
        self.expected_note = expected_note

    def run(self, doc):
        try:
            errors = []
            found = False

            for i, word_range in enumerate(doc.Words):
                word_text = word_range.Text.strip().rstrip(".,;:!?")
                if word_text.lower() == self.keyword.lower():
                    found = True
                    # enumerate(doc.Words) 是 0-based，但 Words.Item(n) 是 从 1 开始的，所以 i + 2 才对应 下一个单词
                    lookahead = doc.Words.Item(i + 2) if i + 2 <= doc.Words.Count else None
                    if lookahead and lookahead.Footnotes.Count > 0:
                        note = lookahead.Footnotes(1).Range.Text.strip()
                        print("📌 Footnote found in next word:", note)
                        if not fuzzy_match(self.expected_note, note, threshold=0.75):
                            errors.append(f"Footnote content may be incorrect: '{note}'")
                        break
                    else:
                        errors.append("'habitat' does not have a footnote.")
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
