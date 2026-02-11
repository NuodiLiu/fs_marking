"""
Test 1 exam configuration migrated to new declarative architecture.

This shows backward compatibility - we can still use old rule classes
while gradually migrating to atomic checks.
"""

from core.rules_old.title_rule import CoverPageTitleRule
from core.rules_old.cover_page_table_rule import CoverPageTableRule
from core.rules_old.margin_rule import MarginRule
from core.rules_old.pagebreak_rule import PageBreakBeforeHeadingRule
from core.rules_old.toc_rule import TableOfContentsRule
from core.rules_old.footer_rule import FooterRule
from core.rules_old.paragraph_indent_rule import ReferencesHangingIndentRule
from core.rules_old.image_rule import ImageRightOfTextRule
from core.rules_old.footnote_rule import FootnoteOnHabitatRule
from core.rules_old.heading_rules.heading1_text_rule import StrictHeading1Rule
from core.rules_old.heading_rules.heading2_text_rule import StrictHeading2Rule
from core.rules_old.multilevel_list_rules.combined_multilevel_rule import CombinedMultilevelListRule
from core.rules_old.style_setting_rules.combined_style_rule import CombinedStyleRule


# Declarative configuration for Test 1
TEST1_CONFIG = {
    "exam": {
        "name": "Test 1 - Australian Ibis Report",
        "code": "test1",
        "total_marks": 13,
        "description": "Original exam with complex formatting requirements"
    },
    
    # Note: We're using old rule classes here for backward compatibility
    # These will be gradually migrated to atomic checks
    "rules": [
        MarginRule(),
        CoverPageTitleRule(),
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
    ],
    
    "output": {
        "template_path": "template/marking_template.xlsx",
        "output_path": "logs/summary.xlsx",
        "sheet_name": "Summary",
        "header_row": 1,
        "zid_row": 2,
        "margin_row": 4,
        "cover_page_row": 5,
        "ToC_row": 6,
        "H1_row": 7,
        "H2_row": 8,
        "styles": 9,
        "picture_row": 10,
        "Habitat_row": 11,
        "Footer_row": 12,
        "Multilist_row": 13,
        "pageBreak_row": 14,
        "Indent_row": 15,
        "total_row": 16,
    }
}