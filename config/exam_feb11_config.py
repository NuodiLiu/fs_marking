# config/cover_page_example_config.py

"""
示例配置：Cover Page检查
对应要求：
1. 第一页有shape，shape里有"Technology in Everyday Life"
2. 第一页有table (3行2列)
3. Table第2列的第2、3行不为空
4. Table应用了非默认样式
"""

COVER_PAGE_CONFIG = {
    "exam": {
        "name": "Cover Page Example",
        "total_marks": 11
    },
    
    "rules": [
        {
            "name": "All Margins 3cm",
            "mark": 1,
            "checks": [
                {"type": "all_margins", "params": {"expected_cm": 3.0}}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "Cover Page",
            "mark": 1,
            "checks": [
                # 检查shape里有要求的文字（隐含检查了有shape）
                {"type": "shape_contains_text", "params": {
                    "page_num": 1,
                    "expected_text": "Technology in Everyday Life",
                    "fuzzy": True,
                    "threshold": 0.8
                }},
                
                # 检查table尺寸（隐含检查了有table）
                {"type": "table_dimension", "params": {
                    "page_num": 1,
                    "table_index": 0,
                    "rows": 3,
                    "cols": 2
                }},
                
                # 检查第2行第2列不为空（学生姓名）
                {"type": "table_cell_not_empty", "params": {
                    "page_num": 1,
                    "table_index": 0,
                    "row": 2,
                    "col": 2
                }},
                
                # 检查第3行第2列不为空（学生ZID）
                {"type": "table_cell_not_empty", "params": {
                    "page_num": 1,
                    "table_index": 0,
                    "row": 3,
                    "col": 2
                }},
                
                {"type": "table_has_style", "params": {
                    "page_num": 1,
                    "table_index": 0,
                    "styles": [
                        "Table Grid"
                    ],
                    "mode": "blacklist"
                }}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "First Page Shadow Border",
            "mark": 1,
            "checks": [
                {"type": "page_border_shadow", "params": {"page_num": 1}}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "Normal Style Settings",
            "mark": 1,
            "checks": [
                {"type": "style_font", "params": {
                    "style_name": "Normal",
                    "font_name": "Arial",
                    "font_size": 14
                }},
                {"type": "style_paragraph_format", "params": {
                    "style_name": "Normal",
                    "line_spacing_rule": "1.5"
                }}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "Heading 1 Settings",
            "mark": 1,
            "checks": [
                {"type": "style_font", "params": {
                    "style_name": "Heading 1",
                    "color": ["not_blue", "not_black"]
                }},
                {"type": "style_paragraph_format", "params": {
                    "style_name": "Heading 1",
                    "space_before": 24
                }},
                {"type": "style_paragraph_format", "params": {
                    "style_name": "Heading 1",
                    "space_after": 10
                }}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "TOC Requirements",
            "mark": 1,
            "checks": [
                {"type": "has_toc", "params": {"page_num": 2}},
                {"type": "toc_on_separate_page", "params": {"page_num": 2}}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "Header Content",
            "mark": 1,
            "checks": [
                {"type": "section_has_header", "params": {"section_index": 1}}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "Section Structure Penalty",
            "mark": -1,
            "checks": [
                {"type": "section_has_header", "params": {"section_index": 1}},
                {"type": "section_has_footer", "params": {"section_index": 1}}
            ],
            "scoring": "at_least",
            "min_pass": 1
        },
        {
            "name": "Footer Before Second Paragraph",
            "mark": 0,
            "checks": [],
            "scoring": "all_or_nothing",
            "placeholder": True
        },
        {
            "name": "Second Paragraph Indent",
            "mark": 1,
            "checks": [
                {"type": "paragraph_format", "params": {
                    "text_start": "Over the past few decades",
                    "left_indent": 1.0,
                    "right_indent": 1.0,
                    "fuzzy": True,
                    "threshold": 0.7
                }}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "Main Points List Format",
            "mark": 1,
            "checks": [
                {"type": "simple_bullet_list", "params": {
                    "text_start": "Mobile Devices",
                    "align_at_left": 0.0,
                    "exclude_bullets": ["\u2022", "\u25E6", "\u25CB"],
                    "fuzzy": True,
                    "threshold": 0.7
                }},
                {"type": "multilevel_list_item", "params": {
                    "text_start": "Mobile Devices",
                    "level": 1,
                    "align_at_left": 0.0,
                    "exclude_bullets": ["\u2022", "\u25E6", "\u25CB"],
                    "fuzzy": True,
                    "threshold": 0.7
                }},
                {"type": "simple_bullet_list", "params": {
                    "text_start": "Online Platforms",
                    "align_at_left": 0.0,
                    "exclude_bullets": ["\u2022", "\u25E6", "\u25CB"],
                    "fuzzy": True,
                    "threshold": 0.7
                }},
                {"type": "multilevel_list_item", "params": {
                    "text_start": "Online Platforms",
                    "level": 1,
                    "align_at_left": 0.0,
                    "exclude_bullets": ["\u2022", "\u25E6", "\u25CB"],
                    "fuzzy": True,
                    "threshold": 0.7
                }}
            ],
            "scoring": "at_least",
            "min_pass": 2
        },
        {
            "name": "Conclusion Section",
            "mark": 1,
            "checks": [
                # Check "Conclusion" appears exactly once in headings
                {"type": "text_occurrence", "params": {
                    "text": "Conclusion",
                    "heading_only": True,
                    "expected_count": 1,
                    "threshold": 0.7
                }},
                # Check last heading is "Conclusion"
                {"type": "last_heading", "params": {
                    "text": "Conclusion",
                    "threshold": 0.7
                }},
                # Check conclusion paragraph appears exactly once in body
                {"type": "text_occurrence", "params": {
                    "text": "Technology in everyday life continues",
                    "body_only": True,
                    "expected_count": 1,
                    "threshold": 0.7
                }},
                # Check last body paragraph starts with conclusion text
                {"type": "last_paragraph", "params": {
                    "text": "Technology in everyday life continues",
                    "threshold": 0.7
                }}
            ],
            "scoring": "all_or_nothing"
        },
        {
            "name": "Textbox",
            "mark": 0,
            "checks": [],
            "scoring": "all_or_nothing",
            "placeholder": True
        },
        {
            "name": "Footnote for Effect",
            "mark": 1,
            "checks": [
                {"type": "footnote", "params": {
                    "keyword": "effect",
                    "expected_text": "Australian Digital Habits Report (2024)",
                    "fuzzy": True,
                    "threshold": 0.75,
                    "case_sensitive": False
                }}
            ],
            "scoring": "all_or_nothing"
        },


    ],
    
    "merge_rules": [
        {
            "display_name": "Style Settings",
            "rules": ["Normal Style Settings", "Heading 1 Settings"],
            "show_details": False
        },
        {
            "display_name": "Header and Structure",
            "rules": ["Header Content", "Section Structure Penalty"],
            "show_details": False
        }
    ],
    
    "output": {
        "excel_summary": True,
        "feedback_files": True
    }
}
