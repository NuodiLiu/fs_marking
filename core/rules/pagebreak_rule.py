from core.rules.base_rule import BaseRule


class PageBreakBeforeHeadingRule(BaseRule):
    def __init__(self, mark=1, keyword="SUMMARY"):
        super().__init__(f"Page Break Before Heading '{keyword}'", mark)
        self.keyword = keyword

    def run(self, doc):
        try:
            paragraphs = doc.Paragraphs
            errors = []
            found = False

            for i in range(len(paragraphs) - 1, -1, -1):
                para = paragraphs.Item(i + 1)
                text = para.Range.Text.strip()

                if text.upper() == self.keyword.upper():
                    found = True

                    if i == 0:
                        raise Exception(f"'{self.keyword}' is the first paragraph; this is invalid structure.")
                    else:
                        prev_para = paragraphs.Item(i)
                        # 3 is wdActiveEndPageNumber
                        summary_page = para.Range.Information(3)
                        prev_page = prev_para.Range.Information(3)
                        if summary_page == prev_page:
                            errors.append(f"'{self.keyword}' is not on a new page (no page break before it).")
                    break

            if not found:
                raise Exception(f"Could not find heading '{self.keyword}' in the document.")

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
                "errors": [f"Error checking pagebreak before '{self.keyword}': {str(e)}"],
                "needs_review": True
            }