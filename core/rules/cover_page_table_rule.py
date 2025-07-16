from core.rules.base_rule import BaseRule
from core.utils.utils import validate_zid
from win32com.client import constants

class CoverPageTableRule(BaseRule):
    def __init__(self, mark=1):
        super().__init__("Cover Page Table Check", mark)

    def run(self, doc):
        try:
            word_art_bottom = 0
            # 找到第一页上的 WordArt 的底部位置
            for shape in doc.Shapes:
                if shape.Type == constants.msoTextEffect:
                    page_num = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                    if page_num == 1:
                        word_art_bottom = max(word_art_bottom, shape.Top + shape.Height)

            valid_styles = [
                "Grid Table 4 - Accent 1",
                "Grid Table 1 - Accent 4"
            ]

            for table in doc.Tables:
                page_num = table.Range.Information(constants.wdActiveEndPageNumber)
                if page_num != 1:
                    continue

                errors = []

                # ✅ 检查大小
                if table.Rows.Count != 2 or table.Columns.Count != 2:
                    errors.append(f"Table size is {table.Rows.Count}x{table.Columns.Count}, expected 2x2.")
                
                # Check that:
                # - Cell(2,1) contains non-empty name string
                # - Cell(2,2) contains a valid ZID using validate_zid()
                try:
                    name_text = table.Cell(2, 1).Range.Text.strip().replace('\r', '').replace('\x07', '')
                    if not name_text:
                        errors.append("Name in cell (2,1) is empty.")
                except Exception:
                    errors.append("Failed to extract name from cell (2,1).")

                try:
                    zid_text = table.Cell(2, 2).Range.Text.strip().replace('\r', '').replace('\x07', '')
                    if not validate_zid(zid_text):
                        errors.append(f"ZID '{zid_text}' is invalid.")
                except Exception:
                    errors.append("Failed to extract ZID from cell (2,2).")

                # ✅ 检查样式
                style_name = table.Style.NameLocal
                if style_name not in valid_styles:
                    errors.append(f"Table style is '{style_name}', expected one of {valid_styles}.")

                # ✅ 检查位置是否在 WordArt 之后（下方）
                try:
                    if table.Top < word_art_bottom:
                        errors.append("Table appears above WordArt.")
                except Exception:
                    # 有些 table.Top 在某些 Word 版本中可能不可访问
                    pass

                if not errors:
                    return {
                        "name": self.name,
                        "mark": self.mark,
                        "errors": [],
                        "needs_review": False
                    }
                else:
                    return {
                        "name": self.name,
                        "mark": 0,
                        "errors": errors,
                        "needs_review": False
                    }

            return {
                "name": self.name,
                "mark": 0,
                "errors": ["No table found on the first page."],
                "needs_review": False
            }

        except Exception as e:
            return {
                "name": self.name,
                "mark": 0,
                "errors": [f"Error during cover page table check: {str(e)}"],
                "needs_review": True
            }
