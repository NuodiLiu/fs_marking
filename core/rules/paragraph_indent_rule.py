from core.rules.base_rule import BaseRule
from core.utils.utils import fuzzy_match
from core.utils.paragraph_utils import check_paragraph_indent

class ReferencesHangingIndentRule(BaseRule):
    EXPECTED_LINES = [
        "Australian White Ibis - BirdLife Australia. Retrieved from https://birdlife.org.au/bird-profiles/australian-white-ibis/",
        "Urban Pest or Aussie Hero? Changing Media Representations of the Australian White Ibis. MDPI. Retrieved from https://www.mdpi.com/2076-2615/14/22/3251",
        "Australian White Ibis - The Australian Museum. Retrieved from https://australian.museum/learn",
    ]

    def __init__(self, mark=1):
        super().__init__("Hanging indent for References", mark)

    def run(self, doc):
        try:
            paragraphs = doc.Paragraphs
            matched_paras = []
            expected_idx = 0

            for para in paragraphs:
                text = para.Range.Text.strip()
                if not text:
                    continue

                if fuzzy_match(self.EXPECTED_LINES[expected_idx], text, threshold=0.85):
                    matched_paras.append(para)
                    expected_idx += 1

                    if expected_idx >= len(self.EXPECTED_LINES):
                        break
                elif expected_idx > 0:
                    break  # 顺序被打乱了，终止匹配

            if len(matched_paras) != 3:
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": [f"Could not match all 3 reference entries. Only found {len(matched_paras)}."],
                    "needs_review": True
                }

            failed = []
            for i, para in enumerate(matched_paras):
                if not check_paragraph_indent(para, indent_type="hanging", expected_indent=2):
                    failed.append(f"Line {i+1} does not have correct 2cm hanging indent.")

            if failed:
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": failed,
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
                "errors": [f"Error in hanging indent check: {str(e)}"],
                "needs_review": True
            }
