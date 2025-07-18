# core/rules/image_rules/image_right_of_text_rule.py

from core.rules.base_rule import BaseRule
from core.utils.utils import fuzzy_match
from win32com.client import constants
import difflib

def inspect_images(doc):
    print("🔍 Floating Shapes (doc.Shapes):")
    for i, shape in enumerate(doc.Shapes, 1):
        print(f"  ➤ Shape {i}: Type = {shape.Type}")
        if shape.Type == constants.msoPicture:
            anchor_text = shape.Anchor.Text.strip()
            print(f"    ✅ Floating Picture")
            print(f"    ↪ Anchored to text: {repr(anchor_text[:50])}")
            print(f"    ↪ Position: Left={shape.Left}, Top={shape.Top}")
            print(f"    ↪ Size: Width={shape.Width}, Height={shape.Height}")
            print(f"    ↪ Wrap Type: {shape.WrapFormat.Type}")
        else:
            print(f"    ❌ Not a picture (type={shape.Type})")

    print("\n🔍 Inline Shapes (doc.InlineShapes):")
    for i, inline_shape in enumerate(doc.InlineShapes, 1):
        print(f"  ➤ InlineShape {i}: Type = {inline_shape.Type}")
        if inline_shape.Type == constants.wdInlineShapePicture:
            text_snippet = inline_shape.Range.Text.strip()
            print(f"    ✅ Inline Picture")
            print(f"    ↪ Appears with text: {repr(text_snippet[:50])}")
            print(f"    ↪ Size: Width={inline_shape.Width}, Height={inline_shape.Height}")
        else:
            print(f"    ❌ Not a picture (type={inline_shape.Type})")

class ImageRightOfTextRule(BaseRule):
    def __init__(self, mark=1):
        super().__init__("Image Right of Text Rule", mark)
        self.reference_text = (
            "The Australian White Ibis (Threskiornis molucca) is a distinctive bird species native to Australia. The ibis has undergone significant changes in its habitat and role within Australian society."
        )

    def run(self, doc):
        try:
            # Step 1: Fuzzy find reference paragraph
            #! inspect_images(doc)
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

            page_width = doc.PageSetup.PageWidth

            found = False
            errors = []

            for shape in doc.Shapes:
                # skip shapes that are not a picture
                if shape.Type != constants.msoPicture:
                    continue

                shape_top = shape.Top
                shape_bottom = shape.Top + shape.Height
                
                # 获取段落位置
                para_top_raw = best_para.Range.Information(constants.wdVerticalPositionRelativeToPage)
                margin_top = doc.PageSetup.TopMargin
                para_top = para_top_raw - margin_top

                # 设定允许段落略高于图片顶部的容差（如20 pt）
                tolerance = 20

                # 判断段落顶部是否在图片内部或略高一点
                if not (shape_top - tolerance <= para_top <= shape_bottom):
                    continue

                found = True  # there's an image anchored to the correct paragraph

                # 1. Tight wrap check
                if shape.WrapFormat.Type != constants.wdWrapTight:
                    errors.append("Picture wrap is not set to 'Tight'.")

                # 2. Horizontal position: must be right of paragraph (center is > middle of page)
                center_x = shape.Left + shape.Width / 2
                if shape.Left == -999996: # this constant means right aligned
                    pass
                elif center_x <= page_width / 2:
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
