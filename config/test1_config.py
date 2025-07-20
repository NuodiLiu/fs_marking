from core.rules.cover_page_table_rule import CoverPageTableRule
from core.rules.footer_rule import FooterRule
from core.rules.footnote_rule import FootnoteOnHabitatRule
from core.rules.heading_rules.heading1_text_rule import StrictHeading1Rule
from core.rules.heading_rules.heading2_text_rule import StrictHeading2Rule
from core.rules.image_rule import ImageRightOfTextRule
from core.rules.margin_rule import MarginRule
from core.rules.multilevel_list_rules.combined_multilevel_rule import \
    CombinedMultilevelListRule
from core.rules.pagebreak_rule import PageBreakBeforeHeadingRule
from core.rules.paragraph_indent_rule import ReferencesHangingIndentRule
from core.rules.style_setting_rules.combined_style_rule import \
    CombinedStyleRule
from core.rules.title_rule import CoverPageTitleRule
from core.rules.toc_rule import TableOfContentsRule


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

test1_config = Test1Config()