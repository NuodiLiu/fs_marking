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
        "total_marks": 3
    },
    
    "rules": [
        {
            "name": "All Margins 3cm",
            "max_marks": 1,
            "checks": [
                {"type": "all_margins", "params": {"expected_cm": 3.0}}
            ],
            "scoring_mode": "all_or_nothing"
        },
        {
            "name": "Cover Page",
            "max_marks": 1,
            "checks": [
                # 检查第一页有shape
                {"type": "page_has_shape", "params": {"page_num": 1}},
                
                # 检查shape里有要求的文字（fuzzy match）
                {"type": "shape_contains_text", "params": {
                    "page_num": 1,
                    "expected_text": "Technology in Everyday Life",
                    "fuzzy": True,
                    "threshold": 0.8
                }},
                
                # 检查第一页有table
                {"type": "page_has_table", "params": {"page_num": 1}},
                
                # 检查table尺寸为3行2列
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
                
                # 检查table样式不是默认样式（白名单模式）
                {"type": "table_has_style", "params": {
                    "page_num": 1,
                    "table_index": 0,
                    "valid_styles": [
                        "Grid Table 4 - Accent 1",
                        "Grid Table 1 - Accent 4",
                        "Grid Table 5 Dark - Accent 1",
                        # 可以添加更多允许的样式
                    ],
                    "mode": "whitelist"
                }}
            ],
            "scoring_mode": "all_or_nothing"  # 所有checks通过才得1分
        },
        {
            "name": "First Page Shadow Border",
            "max_marks": 1,
            "checks": [
                {"type": "page_border_shadow", "params": {"page_num": 1}}
            ],
            "scoring_mode": "all_or_nothing"
        }
    ],
    
    "output": {
        "excel_summary": True,
        "feedback_files": True
    }
}


# 或者使用黑名单模式（不是默认样式即可）
COVER_PAGE_CONFIG_BLACKLIST = {
    "exam": {
        "name": "Cover Page Example (Blacklist)",
        "total_marks": 1
    },
    
    "rules": [
        {
            "name": "Cover Page",
            "max_marks": 1,
            "checks": [
                {"type": "page_has_shape", "params": {"page_num": 1}},
                {"type": "shape_contains_text", "params": {
                    "page_num": 1,
                    "expected_text": "Technology in Everyday Life",
                    "fuzzy": True
                }},
                {"type": "page_has_table", "params": {"page_num": 1}},
                {"type": "table_dimension", "params": {
                    "page_num": 1,
                    "rows": 3,
                    "cols": 2
                }},
                {"type": "table_cell_not_empty", "params": {
                    "page_num": 1,
                    "row": 2,
                    "col": 2
                }},
                {"type": "table_cell_not_empty", "params": {
                    "page_num": 1,
                    "row": 3,
                    "col": 2
                }},
                
                # 黑名单模式：只要不是默认样式就行
                {"type": "table_has_style", "params": {
                    "page_num": 1,
                    "valid_styles": ["Table Grid", "Normal Table", "Table Normal"],
                    "mode": "blacklist"
                }}
            ],
            "scoring_mode": "all_or_nothing"
        }
    ],
    
    "output": {
        "excel_summary": True
    }
}
