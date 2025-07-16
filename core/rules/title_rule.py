from core.rules.base_rule import BaseRule
from core.utils.utils import fuzzy_match
from win32com.client import constants # type: ignore

class CoverPageTitleRule(BaseRule):
    def __init__(self, mark=1, expected_title="The Australian Ibis"):
        super().__init__("Cover Page Title Check", mark)
        self.expected_title = expected_title

    def run(self, doc):
        try:
            page_width = doc.PageSetup.PageWidth
            center_x = page_width / 2
            tolerance = 50  # pt

            found_wordart = False
            matched_text = False
            centered_position = False

            for shape in doc.Shapes:
                page_num = shape.Anchor.Information(constants.wdActiveEndPageNumber)

                if page_num == 1 and shape.Type == constants.msoTextEffect:
                    found_wordart = True
                    text = shape.TextEffect.Text.strip()

                    matched_text = fuzzy_match(self.expected_title, text, threshold=0.8)
                    centered_position = abs((shape.Left + shape.Width / 2) - center_x) <= tolerance

                    if matched_text and centered_position:
                        return {
                            "name": self.name,
                            "mark": self.mark,
                            "errors": [],
                            "needs_review": False
                        }

            errors = []
            if not found_wordart:
                errors.append("No WordArt object found on the first page.")
            elif not matched_text:
                errors.append("WordArt text does not match the expected title.")
            if found_wordart and not centered_position:
                errors.append("WordArt title is not approximately centered on the page.")

            return {
                "name": self.name,
                "mark": 0,
                "errors": errors,
                "needs_review": False
            }

        except Exception as e:
            return {
                "name": self.name,
                "mark": 0,
                "errors": [f"Error during cover title check: {str(e)}"],
                "needs_review": True
            }
