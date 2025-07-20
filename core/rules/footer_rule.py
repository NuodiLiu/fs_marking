from core.rules.base_rule import BaseRule
from win32com.client import constants
import re

class FooterRule(BaseRule):
    def __init__(self, mark=2):
        super().__init__("Footer Rule", mark)
        self.page_width = 0

    def _extract_footer_candidate_line(self, footer_range):
        paragraphs = footer_range.Paragraphs
        candidate_ranges = [p.Range for p in paragraphs if p.Range.Text.strip()]

        # 优先选中含 ZID 的行
        for rng in candidate_ranges:
            if re.search(r"z\d{7}", rng.Text, re.IGNORECASE):
                return rng

        # 其次选中包含“page”或“name”的行
        for rng in candidate_ranges:
            text = rng.Text.lower()
            if "page" in text or "name" in text:
                return rng

        # fallback：第一非空段落
        return candidate_ranges[0] if candidate_ranges else None


    def _split_footer_parts(self, footer_range):
        line_range = self._extract_footer_candidate_line(footer_range)
        if not line_range:
            return ("", "", "")

        text = line_range.Text.strip()
        parts = [p.strip() for p in text.split('\t')]
        while len(parts) < 3:
            parts.append("")

        # ✅ 默认 parts = [left, center, right]
        left, center, right = parts[:3]

        # ✅ 检查是否有 PAGE 字段对象落在这个 range 内
        page_field = None
        for field in footer_range.Fields:
            if field.Type == constants.wdFieldPage and line_range.Start <= field.Result.Start <= line_range.End:
                page_field = field
                break

        if page_field:
            try:
                x_pos = page_field.Result.Information(constants.wdHorizontalPositionRelativeToPage)
                page_width = self.page_width  # in points
                
                if x_pos < page_width / 3:
                    left = "PAGE"
                elif x_pos < 2 * page_width / 3:
                    center = "PAGE"
                else:
                    right = "PAGE"
            except Exception as e:
                # fallback 逻辑（可选）
                print("⚠️ Failed to get PAGE position:", e)

        return (left, center, right)

    def _validate_footer_parts(self, left: str, center: str, right: str) -> list[str]:
        errors = []
        if not left or not center or not right:
            errors.append("Some part of footer is missing.")

        if "PAGE" not in (left, center, right):
            errors.append("Missing page number field")

        return errors

    def _footer_has_valid_parts(self, footer_range):
        line = self._extract_footer_candidate_line(footer_range)

        if not line:
            return ["Footer is empty"], ("", "", "")

        left, center, right = self._split_footer_parts(line)

        print(left, ' ', center, ' ', right)
        errors = self._validate_footer_parts(left, center, right)

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
            self.page_width = doc.PageSetup.PageWidth
            section_count = doc.Sections.Count
            errors = []
            print(f"section_cout: {section_count}")
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
                if err:
                    return {
                        "name": self.name,
                        "mark": 0,
                        "errors": ["Only one section and footer exists but not valid, it may contains confusing footer"] + err,
                        "needs_review": True
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

            sec1_contains_footer = False
            sec2_contains_footer = True
            
            if sec1_footer.Range.Text.strip():
                errors.append("Section 1 should not contain footer")
                sec1_contains_footer = True

            if not sec2_footer.Range.Text.strip():
                errors.append("Section 2 is missing footer")
                sec2_contains_footer = False
            
            if not sec2_contains_footer:
                # 如果 section 2 没有 footer，无论 section 1 是否有，都是错误
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": errors,
                    "needs_review": False
                }

            content_errors, _ = self._footer_has_valid_parts(sec2_footer.Range)
            if content_errors:
                errors.extend(content_errors)
            footer_valid = len(content_errors) == 0

            if not footer_valid:
                return {
                    "name": self.name,
                    "mark": 0,
                    "errors": errors,
                    "needs_review": True
                }

            if sec1_contains_footer:
                return {
                    "name": self.name,
                    "mark": 1,
                    "errors": errors,
                    "needs_review": False
                }

            return {
                "name": self.name,
                "mark": 2,
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
