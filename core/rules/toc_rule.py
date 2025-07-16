from core.rules.base_rule import BaseRule
from win32com.client import constants

class TableOfContentsRule(BaseRule):
    def __init__(self):
        super().__init__("Table of Contents Rule", 1)

    def run(self, doc):
        try:
            found_toc = False
            extra_text_on_page_2 = False

            for field in doc.Fields:
                if field.Type == constants.wdFieldTOC:
                    page_number = field.Result.Information(constants.wdActiveEndPageNumber)
                    if page_number == 2:
                        found_toc = True

                        # 检查第 2 页是否还有别的内容
                        page_2_start = field.Result.Start
                        page_2_end = field.Result.End

                        page_2_range = doc.Range(page_2_start, page_2_end)

                        # 检查是否除了TOC还有其他段落或文字
                        for para in page_2_range.Paragraphs:
                            text = para.Range.Text.strip()
                            if text and "contents" not in text.lower():
                                extra_text_on_page_2 = True
                                break
                        break  # 只看第一个TOC

            if not found_toc:
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": ["No TOC found on page 2"],
                    "needs_review": False
                }

            if extra_text_on_page_2:
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": ["Page 2 contains content other than TOC"],
                    "needs_review": False
                }

            return {
                "name": self.name,
                "mark": 1,
                "errors": [],
                "needs_review": False
            }

        except Exception as e:
            return {
                "name": self.name,
                "mark": 0,
                "errors": [f"Error checking TOC on page 2: {str(e)}"],
                "needs_review": True
            }
