# core/writers/excel_writer.py

import os
import pandas as pd

class ExcelSummaryWriter:
    def __init__(self, summary_path="logs/summary.xlsx"):
        self.summary_path = summary_path
        self.rows = []  # 缓存所有学生结果

    def write(self, zid: str, result: dict):
        row = {"ZID": zid, "Total": result["total"]}
        for rule in result["results"]:
            row[rule["name"]] = rule["mark"]
        self.rows.append(row)

    def save(self):
        df = pd.DataFrame(self.rows)
        df.to_excel(self.summary_path, index=False)
