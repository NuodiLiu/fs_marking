"""
Test 2 exam configuration using pure declarative style.

This demonstrates the new modular architecture - configuration is pure data,
no need to instantiate classes or use builders.
"""

# Pure declarative configuration - just data!
TEST2_CONFIG = {
    # Exam metadata
    "exam": {
        "name": "Test 2 - Digital Report",
        "code": "test2",
        "total_marks": 3,
        "description": "Formatting assessment for digital report submission"
    },
    
    # Rules definition - pure data structures
    "rules": [
        # Rule 1: All margins must be 3cm (1 point)
        {
            "name": "All Margins 3cm",
            "mark": 1,
            "checks": [
                ("margin", {"side": "Top", "expected_cm": 3.0}),
                ("margin", {"side": "Bottom", "expected_cm": 3.0}),
                ("margin", {"side": "Left", "expected_cm": 3.0}),
                ("margin", {"side": "Right", "expected_cm": 3.0}),
            ],
            "scoring": "all_or_nothing"
        },
        
        # Rule 2: Header must have any string (1 point)
        # Assuming main content starts at section 3
        {
            "name": "Header Has Content",
            "mark": 1,
            "checks": [
                ("header_content", {"section_index": 3, "min_length": 1}),
            ],
            "scoring": "all_or_nothing"
        },
        
        # Rule 3: Cover page and TOC should not have header/footer
        # This is a DEDUCTION rule (negative mark)
        {
            "name": "Cover & TOC: No Header/Footer",
            "mark": -1,
            "checks": [
                ("section_no_header", {"section_index": 1}),
                ("section_no_footer", {"section_index": 1}),
                ("section_no_header", {"section_index": 2}),
                ("section_no_footer", {"section_index": 2}),
            ],
            "scoring": "all_or_nothing"
        },
    ],
    
    # Output configuration
    "output": {
        "template_path": "template/marking_template.xlsx",
        "output_path": "logs/test2_summary.xlsx",
        "sheet_name": "Summary",
        "header_row": 1,
        "zid_row": 2,
        "margin_row": 4,
        "header_content_row": 5,
        "cover_toc_row": 6,
        "total_row": 7,
    }
}


class Test2ConfigAlternative:
    """
    Alternative configuration using convenience functions.
    
    This is more concise for common patterns.
    """
    
    def __init__(self):
        builder = RuleBuilder()
        
        # Using convenience functions
        builder.add_existing_rule(margin_rule(3.0, mark=1))
        builder.add_existing_rule(header_exists_rule(section_index=3, mark=1))
        builder.add_existing_rule(no_header_footer_rule([1, 2], mark=-1))
        
        self.RULES = builder.build()
        
        self.OUTPUT = {
            "template_path": "template/marking_template.xlsx",
            "output_path": "logs/test2_summary.xlsx",
            "sheet_name": "Summary",
            "header_row": 1,
            "zid_row": 2,
            "margin_row": 4,
            "header_row_mark": 5,
            "cover_toc_row": 6,
            "total_row": 7,
        }


class Test1ConfigModular:
    """
    Test 1 configuration migrated to use modular architecture.
    
    This shows how to migrate existing rules while keeping backward compatibility.
    """
    
    def __init__(self):
        builder = RuleBuilder()
        
        # Modular margin rule (replaces MarginRule)
        builder.add_rule(
            name="All Margins 2.5cm",
            mark=1,
            checks=[
                ("all_margins", {"expected_cm": 2.5}),  # Using the convenience check
            ],
            scoring="all_or_nothing"
        )
        
        # Keep existing complex rules for now
        # Will be migrated gradually
        from core.rules.title_rule import CoverPageTitleRule
        from core.rules.cover_page_table_rule import CoverPageTableRule
        from core.rules.toc_rule import TableOfContentsRule
        from core.rules.footer_rule import FooterRule
        
        builder.add_existing_rule(CoverPageTitleRule())
        builder.add_existing_rule(CoverPageTableRule())
        builder.add_existing_rule(TableOfContentsRule())
        builder.add_existing_rule(FooterRule())
        
        self.RULES = builder.build()
        
        self.OUTPUT = {
            "template_path": "template/marking_template.xlsx",
            "output_path": "logs/summary.xlsx",
            "sheet_name": "Summary",
            "header_row": 1,
            "zid_row": 2,
            "margin_row": 4,
            "cover_page_row": 5,
            "ToC_row": 6,
            "Footer_row": 12,
            "total_row": 16,
        }


# Default config for testing
test2_config = Test2Config()
