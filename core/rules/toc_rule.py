from core.rules.base_rule import BaseRule
from win32com.client import constants

class TableOfContentsRule(BaseRule):
    def __init__(self):
        super().__init__("Table of Contents Rule", 1)

    def run(self, doc):
        try:
            # 1. 找到第 2 页的第一个 TOC 字段
            toc_field = None
            for field in doc.Fields:
                try:
                    if (field.Type == constants.wdFieldTOC and
                        field.Result.Information(constants.wdActiveEndPageNumber) == 2):
                        toc_field = field
                        break
                except:
                    continue

            if toc_field is None:
                return {"name": self.name, "mark": 0,
                        "errors": ["No TOC found on page 2"], "needs_review": False}

            toc_range = toc_field.Result
            toc_text = toc_range.Text  # TOC 条目部分文本

            # 2. 定位第 2 页的开始/结束
            start = end = None
            for para in doc.Paragraphs:
                if para.Range.Information(constants.wdActiveEndPageNumber) == 2:
                    if start is None:
                        start = para.Range.Start
                    end = para.Range.End
            if start is None:
                return {"name": self.name, "mark": 0,
                        "errors": ["Could not locate any content on page 2."],
                        "needs_review": True}

            page2_range = doc.Range(start, end)
            page2_text = page2_range.Text

            # 3. 找出紧靠 TOC 之前的段落，通常就是 “Contents” 这样的标题
            header_text = ""
            for para in doc.Paragraphs:
                if (para.Range.Information(constants.wdActiveEndPageNumber) == 2 and
                    para.Range.End < toc_range.Start):
                    header_text = para.Range.Text
                # 一旦遇到第一个 start > toc_range.Start，就可以跳出
                if para.Range.Start > toc_range.Start:
                    break

            # 4. 从整页文本中去掉 header 和 toc_text，再看剩余是否有非空内容
            cleaned = page2_text
            for fragment in (header_text, toc_text):
                if fragment:
                    cleaned = cleaned.replace(fragment, "")

            if cleaned.strip():
                return {"name": self.name, "mark": 0,
                        "errors": ["Page 2 contains content other than TOC"], 
                        "needs_review": False}

            # 全部检查通过
            return {"name": self.name, "mark": 1, "errors": [], "needs_review": False}

        except Exception as e:
            return {"name": self.name, "mark": 0,
                    "errors": [f"Error checking TOC on page 2: {e}"],
                    "needs_review": True}
