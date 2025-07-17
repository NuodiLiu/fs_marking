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

            for i in range(len(paragraphs) - 1, -1, -1):  # backward search to find keyword, avoid matching toc
                para = paragraphs[i]
                text = para.Range.Text.strip()

                if text.upper() == self.keyword.upper():
                    found = True

                    # only check page break
                    if i == 0:
                        raise Exception(f"'{self.keyword}' is the first paragraph; this is invalid structure.")
                    else:
                        prev_para = paragraphs[i - 1]
                        prev_text = prev_para.Range.Text

                        if "\x0c" not in prev_text:
                            errors.append(f"No page break immediately before '{self.keyword}'.")

                    break  # check the last matched keyword

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
