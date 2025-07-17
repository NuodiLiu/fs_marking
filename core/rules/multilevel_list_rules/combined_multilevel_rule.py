from core.rules.base_rule import BaseRule
from core.rules.multilevel_list_rules.multilevel_checkpoints_list import get_all_multilevel_checkpoints
from core.utils.utils import fuzzy_match

class CombinedMultilevelListRule(BaseRule):
    EXPECTED_LINES = [
        "Appearance",
        "White plumage with black head and neck",
        "Long, curved black bill",
        "Wingspan up to 120 cm",
        "Habitat",
        "Originally wetlands and floodplains",
        "Now common in urban parks and cities",
        "Found across eastern and northern Australia",
        "Diet",
        "Natural: insects, frogs, crustaceans",
        "Urban: food scraps, garbage",
    ]

    def __init__(self):
        super().__init__("Multilevel List under Main Points", 1)

    def run(self, doc):
        paragraphs = self._locate_list_paragraphs(doc)  # use fuzzy match to search paragraphs
        checkpoints = get_all_multilevel_checkpoints()

        total_passed = 0
        errors = []

        for checkpoint in checkpoints:
            if checkpoint.run(paragraphs):
                total_passed += 1
            else:
                errors.append(checkpoint.description + (" -- " + checkpoint.error if checkpoint.error else ""))

        mark = 1 if total_passed >= 3 else 0
        return {
            "name": self.name,
            "mark": mark,
            "errors": errors,
            "needs_review": any(c.needs_review for c in checkpoints)
        }

    def _locate_list_paragraphs(self, doc):
        paragraphs = doc.Paragraphs
        matched = []
        expected_idx = 0

        for para in paragraphs:
            text = para.Range.Text.strip()
            if not text:
                continue

            if fuzzy_match(self.EXPECTED_LINES[expected_idx], text):
                matched.append(para)
                expected_idx += 1

                if expected_idx >= len(self.EXPECTED_LINES):
                    break
            elif expected_idx > 0:
                break

        return matched