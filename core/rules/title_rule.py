from core.rules.base_rule import BaseRule
from core.utils.object_utils import font_has_effects
from core.utils.utils import fuzzy_match
from win32com.client import constants  # type: ignore

class CoverPageTitleRule(BaseRule):
    def __init__(self, mark=1, expected_title="The Australian Ibis"):
        super().__init__("Cover Page Title Check", mark)
        self.expected_title = expected_title

    def run(self, doc):
        try:
            page_width = doc.PageSetup.PageWidth
            center_x = page_width / 2
            tolerance = 50  # pt
            errors = []

            for i, shape in enumerate(doc.Shapes):
                try:
                    page_num = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                    if page_num != 1:
                        continue

                    if shape.Type != constants.msoTextBox:
                        continue

                    try:
                        text_range = shape.TextFrame.TextRange
                        text = text_range.Text.strip()
                    except:
                        errors.append("Found a textbox on page 1 but could not extract its text.")
                        continue

                    if not text:
                        errors.append("Found a textbox on page 1 but it is empty.")
                        continue

                    font_size = text_range.Font.Size
                    if abs(font_size - 36) > 0.5:
                        errors.append(f"Text box with text '{text}' has font size {font_size}, expected 36.")
                        continue

                    try:
                        font = shape.TextFrame.TextRange.Font
                        has_text_effect = font_has_effects(font)
                    except Exception as e:
                        errors.append(f"Error checking font effects: {e}")
                        has_text_effect = False

                    if not has_text_effect:
                        errors.append(f"Text box with text '{text}' does not use WordArt-style effect.")
                        continue

                    if not fuzzy_match(self.expected_title, text, threshold=0.8):
                        errors.append(f"Text box text '{text}' does not match expected title '{self.expected_title}'.")
                        continue

                    # ✅ Passed all checks
                    return {
                        "name": self.name,
                        "mark": self.mark,
                        "errors": [],
                        "needs_review": False
                    }

                except Exception as shape_error:
                    errors.append(f"Error while checking shape {i+1}: {shape_error}")

            if not errors:
                errors.append("No suitable textbox found on page 1 with required formatting and title.")

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
