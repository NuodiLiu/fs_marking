# core/writers/excel_writer.py

import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from typing import Optional

class ExcelSummaryWriter:
    def __init__(self, summary_path="data/output.xlsx"):
        self.path = summary_path
        self.wb = None  # Initialize to None first
        
        if not os.path.exists(self.path):
            raise FileNotFoundError(
                f"❌ Cannot find Excel summary file at '{self.path}'. "
                "Please make sure the template exists before running."
            )
        
        try:
            self.wb = load_workbook(self.path)
        except Exception as e:
            raise RuntimeError(f"❌ Cannot load Excel file '{self.path}': {e}")

    def write(self, zid: str, result: dict):
        """
        Write the result for the given zid into the correct column of all sheets.
        Rules are written in order starting from row 4.
        
        Args:
            zid: Student ID
            result: {
                "total": int,
                "results": [{"name": str, "mark": int, "errors": list, "needs_review": bool}, ...]
            }
        """
        WARNING_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        
        # Find the column for this ZID in all sheets
        for sheet in self.wb.worksheets:
            row_2 = list(sheet.iter_rows(min_row=2, max_row=2, values_only=True))[0]
            
            for col_idx in range(6, sheet.max_column + 1):
                zid_value = row_2[col_idx - 1]
                if isinstance(zid_value, str) and zid.lower() in zid_value.lower():
                    # Write marks for each rule in order
                    row_offset = 0
                    for rule_result in result["results"]:
                        cell = sheet.cell(row=4 + row_offset, column=col_idx)
                        
                        # Write None for placeholder rules and fill with yellow
                        if rule_result.get("placeholder", False):
                            cell.value = None
                            cell.fill = WARNING_FILL
                            row_offset += 1
                            continue
                        
                        mark = rule_result.get("mark", 0)
                        needs_review = rule_result.get("needs_review", False)
                        
                        if needs_review:
                            cell.value = f"{mark} ⚠️" if mark is not None else "⚠️"
                            cell.fill = WARNING_FILL
                        else:
                            cell.value = mark
                        
                        row_offset += 1
                    
                    break  # Found the column, move to next sheet

    def save(self, output_path: Optional[str] = None):
        """Save the Excel workbook."""
        if self.wb is None:
            print("⚠️  Warning: Excel workbook not loaded, skipping save")
            return
        
        try:
            self.wb.save(output_path or self.path)
        except Exception as e:
            print(f"❌ Failed to save Excel file: {e}")
            raise

    def __del__(self):
        """Destructor to ensure Excel file is saved."""
        if self.wb is not None:
            try:
                self.save()
            except Exception as e:
                # Already printed in save(), don't print again
                pass
