from core.rules.title_rule import CoverPageTitleRule
from core.rules.cover_page_table_rule import CoverPageTableRule
from core.rules.margin_rule import MarginRule
from core.rules.pagebreak_rule import PageBreakBeforeHeadingRule
from core.rules.toc_rule import TableOfContentsRule
from core.rules.footer_rule import FooterRule
from core.rules.paragraph_indent_rule import ReferencesHangingIndentRule
from core.rules.image_rule import ImageRightOfTextRule
from core.rules.footnote_rule import FootnoteOnHabitatRule
from core.rules.heading_rules.heading1_text_rule import StrictHeading1Rule
from core.rules.heading_rules.heading2_text_rule import StrictHeading2Rule
from core.rules.multilevel_list_rules.combined_multilevel_rule import CombinedMultilevelListRule
from core.rules.style_setting_rules.combined_style_rule import CombinedStyleRule

class Test1Config:
    def __init__(self):
        self.RULES = [
            MarginRule(),
            CoverPageTitleRule(), # problem - find center
            CoverPageTableRule(),
            TableOfContentsRule(),
            StrictHeading1Rule(),
            StrictHeading2Rule(),
            CombinedStyleRule(),
            ImageRightOfTextRule(),
            FootnoteOnHabitatRule(),
            FooterRule(),
            CombinedMultilevelListRule(),
            PageBreakBeforeHeadingRule(),
            ReferencesHangingIndentRule(),
        ]

        self.OUTPUT = {
        # The input excel file
        "template_path": "template/marking_template.xlsx",
        # The output result file
        "output_path":   "logs/summary.xlsx",
        "sheet_name":    "Summary",
        "header_row":    1,
        "zid_row":       2,
        "margin_row":    4,
        "cover_page_row":5,
        "ToC_row":       6,
        "H1_row":        7,
        "H2_row":        8,
        "styles":        9,
        "picture_row":   10,
        "Habitat_row":   11,
        "Footer_row":    12,
        "Multilist_row": 13,
        "pageBreak_row": 14,
        "Indent_row":    15,
        "total_row":     16,
}

test1_config = Test1Config()