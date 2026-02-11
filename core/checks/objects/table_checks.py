# core/checks/objects/table_checks.py

from core.checks.base_check import BaseCheck
from win32com.client import constants


class PageHasTableCheck(BaseCheck):
    """检查指定页面是否有table"""
    
    def __init__(self, page_num=1):
        """
        Args:
            page_num: 页码（1-based）
        """
        self.page_num = page_num
    
    def check(self, doc):
        try:
            for table in doc.Tables:
                table_page = table.Range.Information(constants.wdActiveEndPageNumber)
                if table_page == self.page_num:
                    return (True, f"✅ Page {self.page_num} has table")
            
            return (False, f"❌ Page {self.page_num} has no table")
        
        except Exception as e:
            return (False, f"❌ Error checking tables: {str(e)}")


class TableDimensionCheck(BaseCheck):
    """检查指定页面的table尺寸（行数和列数）"""
    
    def __init__(self, page_num=1, table_index=0, rows=None, cols=None):
        """
        Args:
            page_num: 页码（1-based）
            table_index: 表格索引（0-based，同一页面可能有多个表格）
            rows: 期望的行数（None表示不检查）
            cols: 期望的列数（None表示不检查）
        """
        self.page_num = page_num
        self.table_index = table_index
        self.rows = rows
        self.cols = cols
    
    def check(self, doc):
        try:
            tables_on_page = []
            for table in doc.Tables:
                table_page = table.Range.Information(constants.wdActiveEndPageNumber)
                if table_page == self.page_num:
                    tables_on_page.append(table)
            
            if len(tables_on_page) == 0:
                return (False, f"❌ No table found on page {self.page_num}")
            
            if self.table_index >= len(tables_on_page):
                return (False, f"❌ Table index {self.table_index} out of range (found {len(tables_on_page)} tables)")
            
            table = tables_on_page[self.table_index]
            actual_rows = table.Rows.Count
            actual_cols = table.Columns.Count
            
            errors = []
            if self.rows is not None and actual_rows != self.rows:
                errors.append(f"rows: {actual_rows} (expected {self.rows})")
            if self.cols is not None and actual_cols != self.cols:
                errors.append(f"cols: {actual_cols} (expected {self.cols})")
            
            if errors:
                return (False, f"❌ Table dimension mismatch: {', '.join(errors)}")
            
            return (True, f"✅ Table has correct dimensions ({actual_rows}x{actual_cols})")
        
        except Exception as e:
            return (False, f"❌ Error checking table dimensions: {str(e)}")


class TableCellNotEmptyCheck(BaseCheck):
    """检查指定单元格是否非空"""
    
    def __init__(self, page_num=1, table_index=0, row=1, col=1):
        """
        Args:
            page_num: 页码（1-based）
            table_index: 表格索引（0-based）
            row: 行号（1-based）
            col: 列号（1-based）
        """
        self.page_num = page_num
        self.table_index = table_index
        self.row = row
        self.col = col
    
    def check(self, doc):
        try:
            tables_on_page = []
            for table in doc.Tables:
                table_page = table.Range.Information(constants.wdActiveEndPageNumber)
                if table_page == self.page_num:
                    tables_on_page.append(table)
            
            if len(tables_on_page) == 0:
                return (False, f"❌ No table found on page {self.page_num}")
            
            if self.table_index >= len(tables_on_page):
                return (False, f"❌ Table index {self.table_index} out of range")
            
            table = tables_on_page[self.table_index]
            
            try:
                cell_text = table.Cell(self.row, self.col).Range.Text.strip()
                # 移除Word特殊字符
                cell_text = cell_text.replace('\r', '').replace('\x07', '')
                
                if cell_text:
                    return (True, f"✅ Cell ({self.row},{self.col}) is not empty")
                else:
                    return (False, f"❌ Cell ({self.row},{self.col}) is empty")
            
            except Exception as e:
                return (False, f"❌ Cell ({self.row},{self.col}) does not exist: {str(e)}")
        
        except Exception as e:
            return (False, f"❌ Error checking table cell: {str(e)}")


class TableCellContainsTextCheck(BaseCheck):
    """检查指定单元格是否包含指定文字"""
    
    def __init__(self, page_num=1, table_index=0, row=1, col=1, expected_text="", fuzzy=False):
        """
        Args:
            page_num: 页码（1-based）
            table_index: 表格索引（0-based）
            row: 行号（1-based）
            col: 列号（1-based）
            expected_text: 期望的文字内容
            fuzzy: 是否使用包含匹配（True）还是精确匹配（False）
        """
        self.page_num = page_num
        self.table_index = table_index
        self.row = row
        self.col = col
        self.expected_text = expected_text
        self.fuzzy = fuzzy
    
    def check(self, doc):
        try:
            tables_on_page = []
            for table in doc.Tables:
                table_page = table.Range.Information(constants.wdActiveEndPageNumber)
                if table_page == self.page_num:
                    tables_on_page.append(table)
            
            if len(tables_on_page) == 0:
                return (False, f"❌ No table found on page {self.page_num}")
            
            if self.table_index >= len(tables_on_page):
                return (False, f"❌ Table index {self.table_index} out of range")
            
            table = tables_on_page[self.table_index]
            
            try:
                cell_text = table.Cell(self.row, self.col).Range.Text.strip()
                cell_text = cell_text.replace('\r', '').replace('\x07', '')
                
                if self.fuzzy:
                    if self.expected_text.lower() in cell_text.lower():
                        return (True, f"✅ Cell ({self.row},{self.col}) contains '{self.expected_text}'")
                    else:
                        return (False, f"❌ Cell ({self.row},{self.col}) does not contain '{self.expected_text}'")
                else:
                    if cell_text == self.expected_text:
                        return (True, f"✅ Cell ({self.row},{self.col}) matches '{self.expected_text}'")
                    else:
                        return (False, f"❌ Cell ({self.row},{self.col}) is '{cell_text}', expected '{self.expected_text}'")
            
            except Exception as e:
                return (False, f"❌ Cell ({self.row},{self.col}) does not exist: {str(e)}")
        
        except Exception as e:
            return (False, f"❌ Error checking table cell text: {str(e)}")


class TableHasStyleCheck(BaseCheck):
    """检查表格是否应用了指定的Table Design样式"""
    
    def __init__(self, page_num=1, table_index=0, valid_styles=None, mode="whitelist"):
        """
        Args:
            page_num: 页码（1-based）
            table_index: 表格索引（0-based）
            valid_styles: 样式名称列表
            mode: "whitelist"（必须在列表中）或 "blacklist"（不能在列表中）
        """
        self.page_num = page_num
        self.table_index = table_index
        self.valid_styles = valid_styles or []
        self.mode = mode
    
    def check(self, doc):
        try:
            tables_on_page = []
            for table in doc.Tables:
                table_page = table.Range.Information(constants.wdActiveEndPageNumber)
                if table_page == self.page_num:
                    tables_on_page.append(table)
            
            if len(tables_on_page) == 0:
                return (False, f"❌ No table found on page {self.page_num}")
            
            if self.table_index >= len(tables_on_page):
                return (False, f"❌ Table index {self.table_index} out of range")
            
            table = tables_on_page[self.table_index]
            
            try:
                style_name = table.Style.NameLocal
            except:
                style_name = "Unknown"
            
            if self.mode == "whitelist":
                if style_name in self.valid_styles:
                    return (True, f"✅ Table has valid style: '{style_name}'")
                else:
                    return (False, f"❌ Table style is '{style_name}', expected one of {self.valid_styles}")
            
            elif self.mode == "blacklist":
                if style_name not in self.valid_styles:
                    return (True, f"✅ Table has non-default style: '{style_name}'")
                else:
                    return (False, f"❌ Table has default style: '{style_name}'")
            
            return (False, f"❌ Invalid mode: {self.mode}")
        
        except Exception as e:
            return (False, f"❌ Error checking table style: {str(e)}")


class TableCountCheck(BaseCheck):
    """检查指定页面的table数量"""
    
    def __init__(self, page_num=1, expected_count=1, mode="exact"):
        """
        Args:
            page_num: 页码（1-based）
            expected_count: 期望的table数量
            mode: "exact"（精确）, "min"（至少）, "max"（最多）
        """
        self.page_num = page_num
        self.expected_count = expected_count
        self.mode = mode
    
    def check(self, doc):
        try:
            count = 0
            for table in doc.Tables:
                table_page = table.Range.Information(constants.wdActiveEndPageNumber)
                if table_page == self.page_num:
                    count += 1
            
            if self.mode == "exact":
                if count == self.expected_count:
                    return (True, f"✅ Page {self.page_num} has exactly {count} table(s)")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} table(s), expected {self.expected_count}")
            
            elif self.mode == "min":
                if count >= self.expected_count:
                    return (True, f"✅ Page {self.page_num} has {count} table(s) (>= {self.expected_count})")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} table(s), expected at least {self.expected_count}")
            
            elif self.mode == "max":
                if count <= self.expected_count:
                    return (True, f"✅ Page {self.page_num} has {count} table(s) (<= {self.expected_count})")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} table(s), expected at most {self.expected_count}")
            
            return (False, f"❌ Invalid mode: {self.mode}")
        
        except Exception as e:
            return (False, f"❌ Error checking table count: {str(e)}")
