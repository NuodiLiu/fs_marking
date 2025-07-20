# core/writers/excel_writer.py

import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from typing import Optional

class ExcelSummaryWriter:
    def __init__(self, summary_path="logs/summary.xlsx"):
        self.path = summary_path
        self.wb = load_workbook(self.path)
        self.RULE_ROW_MAPPING = [
            0,                      # MarginRule
            (1, 2),                 # CoverPageTitleRule + CoverPageTableRule → merge
            3,                      # TableOfContentsRule
            4,                      # StrictHeading1Rule
            5,                      # StrictHeading2Rule
            6,                      # CombinedStyleRule
            7,                      # ImageRightOfTextRule
            8,                      # FootnoteOnHabitatRule
            9,                      # FooterRule
            10,                     # CombinedMultilevelListRule
            11                      # PageBreakBeforeHeadingRule / ReferencesHangingIndentRule
        ]
    def _convert_result_to_excel_marks(self, result: dict) -> list[int]:
    """
    Convert 13-rule result into 12-row mark list for Excel, with merge logic applied.
    """
        results = result["results"]
        marks = [None] * 12

        for i, rule_map in enumerate(self.RULE_ROW_MAPPING):
        if isinstance(rule_map, tuple):
            combined_mark = sum(results[idx]["mark"] for idx in rule_map)
            marks[i] = combined_mark
        else:
            marks[i] = results[rule_map]["mark"]
        return marks

    def write(self, zid: str, result: dict):
    """
    Write the result for the given zid into the correct column of all sheets.
    """
        marks = self._convert_result_to_excel_marks(result)
        total = result.get("total")
        
        for sheet in self.wb.worksheets:
            row_2 = list(sheet.iter_rows(min_row=2, max_row=2, values_only=True))[0]
            for col_idx in range(6, sheet.max_column + 1):
                zid_value = row_2[col_idx - 1]
                if isinstance(zid_value, str) and zid_value == zid:
                    start_row = 4
                    for offset, mark in enumerate(marks):
                        sheet.cell(row=start_row+offset, column=col_idx, value=mark)
                    sheet.cell(row=16, column=col_idx, value=total)
                    break
    def save(self):
        directory = os.path.dirname(self.path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        self.wb.save(self.path)

