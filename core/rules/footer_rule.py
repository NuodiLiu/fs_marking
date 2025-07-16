from core.rules.base_rule import BaseRule
from win32com.client import constants


class FooterRule(BaseRule):
    def __init__(self, mark=2):
        super().__init__("Footer Rule", mark)

    def _footer_has_valid_parts(self, footer_range):
        paragraphs = footer_range.Paragraphs
        left = paragraphs(1).Range.Text.strip() if paragraphs.Count >= 1 else ""
        center = paragraphs(2).Range.Text.strip() if paragraphs.Count >= 2 else ""
        right = paragraphs(3).Range.Text.strip() if paragraphs.Count >= 3 else ""

        errors = []

        if not left:
            errors.append("Footer left part (Name) is empty")
        if not center:
            errors.append("Footer center part (Page number) is empty")
        if not right:
            errors.append("Footer right part (ZID) is empty")

        return errors, (left, center, right)

    def _contains_page_number_field(self, footer_range):
        try:
            for field in footer_range.Fields:
                if field.Type == constants.wdFieldPage:
                    return True
            return False
        except Exception:
            return False

    def run(self, doc):
        try:
            section_count = doc.Sections.Count
            errors = []

            if section_count < 2:
                section = doc.Sections(1)
                footer = section.Footers(constants.wdHeaderFooterPrimary)

                if not footer.Range.Text.strip():
                    return {
                        "name": self.name,
                        "mark": 0,
                        "errors": ["Only one section and footer is missing"],
                        "needs_review": False
                    }

                err, _ = self._footer_has_valid_parts(footer.Range)
                if err or not self._contains_page_number_field(footer.Range):
                    return {
                        "name": self.name,
                        "mark": 0,
                        "errors": ["Only one section and footer exists but not valid"] + err,
                        "needs_review": False
                    }

                return {
                    "name": self.name,
                    "mark": 1,
                    "errors": ["Only one section, but footer is valid"],
                    "needs_review": False
                }

            # section 1 and 2 exist
            sec1_footer = doc.Sections(1).Footers(constants.wdHeaderFooterPrimary)
            sec2_footer = doc.Sections(2).Footers(constants.wdHeaderFooterPrimary)

            mark = 0

            if sec1_footer.Range.Text.strip():
                errors.append("Section 1 should not contain footer")

            if not sec2_footer.Range.Text.strip():
                errors.append("Section 2 is missing footer")
                return {
                    "name": self.name,
                    "mark": 1,
                    "errors": errors,
                    "needs_review": False
                }

            content_errors, _ = self._footer_has_valid_parts(sec2_footer.Range)
            if content_errors:
                errors.extend(content_errors)

            if not self._contains_page_number_field(sec2_footer.Range):
                errors.append("Section 2 footer is missing auto page number field")

            if not errors:
                mark = 2
            else:
                mark = 1

            return {
                "name": self.name,
                "mark": mark,
                "errors": errors,
                "needs_review": False
            }

        except Exception as e:
            return {
                "name": self.name,
                "mark": 0,
                "errors": [f"Exception occurred during footer check: {str(e)}"],
                "needs_review": True
            }
