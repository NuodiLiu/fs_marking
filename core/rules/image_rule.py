# core/rules/image_rules/image_right_of_text_rule.py

from core.rules.base_rule import BaseRule
from core.utils.utils import fuzzy_match
from win32com.client import constants
import difflib

class ImageRightOfTextRule(BaseRule):
    def __init__(self, mark=1):
        super().__init__("Image Right of Text Rule", mark)
        self.reference_text = (
            "The Australian White Ibis (Threskiornis molucca) is a distinctive bird species native to Australia. "
            "The ibis has undergone significant changes in its habitat and role within Australian society."
        )

    def run(self, doc):
        try:
            # Step 1: Fuzzy find reference paragraph
            best_para = None
            best_score = 0
            for para in doc.Paragraphs:
                text = para.Range.Text.strip()
                score = difflib.SequenceMatcher(None, self.reference_text.lower(), text.lower()).ratio()
                if score > best_score:
                    best_score = score
                    best_para = para
            if best_score < 0.8 or not best_para:
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": ["Failed to find reference paragraph using fuzzy match."],
                    "needs_review": True
                }

            para_range = best_para.Range
            para_start = para_range.Start
            para_end = para_range.End
            page_width = doc.PageSetup.PageWidth

            found = False
            errors = []

            for shape in doc.Shapes:
                # skip shapes that are not a picture
                if shape.Type != constants.msoPicture:
                    continue

                anchor_start = shape.Anchor.Start
                if not (para_start <= anchor_start <= para_end):
                    continue  # only consider images anchored to the reference paragraph

                found = True  # there's an image anchored to the correct paragraph

                # 1. Tight wrap check
                if shape.WrapFormat.Type != constants.wdWrapTight:
                    errors.append("Picture wrap is not set to 'Tight'.")

                # 2. Horizontal position: must be right of paragraph (center is > middle of page)
                center_x = shape.Left + shape.Width / 2
                if center_x <= page_width / 2:
                    errors.append("Picture is not positioned to the right side of the page.")

                # 3. Size check: width < 2/3 page width
                if shape.Width > page_width * 2 / 3:
                    errors.append("Picture width exceeds 2/3 of the page width.")

                # 4. Height should not exceed width (prevent tall vertical layout)
                if shape.Height > shape.Width:
                    errors.append("Picture height is larger than its width, may look disproportionate.")

                break  # Only consider the first matching image

            if not found:
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": ["No picture found anchored to the reference paragraph."],
                    "needs_review": False
                }

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
                "errors": [f"Error during image placement check: {str(e)}"],
                "needs_review": True
            }
