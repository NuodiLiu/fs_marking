# core/utils/section_utils.py

from win32com.client import constants


def get_section_for_page(doc, page_num):
    """
    获取指定页面所属的section及其在section中的位置
    
    Args:
        doc: Word document对象
        page_num: 页码（1-based）
    
    Returns:
        tuple: (section对象, is_first_page_in_section布尔值)
        如果找不到返回 (None, False)
    """
    try:
        for section in doc.Sections:
            # 获取section的起始和结束页码
            section_start = section.Range.Information(constants.wdActiveEndPageNumber)
            section_end_range = section.Range
            section_end_range.Collapse(constants.wdCollapseEnd)
            section_end = section_end_range.Information(constants.wdActiveEndPageNumber)
            
            # 判断page_num是否在这个section范围内
            if section_start <= page_num <= section_end:
                is_first_page = (page_num == section_start)
                return (section, is_first_page)
        
        return (None, False)
    
    except Exception as e:
        return (None, False)


def is_border_enabled_for_page(section, is_first_page):
    """
    判断指定页面是否启用了边框
    
    根据section的EnableFirstPageInSection和EnableOtherPagesInSection属性判断
    
    Args:
        section: Section对象
        is_first_page: 是否为section的第一页
    
    Returns:
        bool: 该页是否启用边框
    """
    try:
        borders = section.Borders
        
        if is_first_page:
            # 第一页：检查EnableFirstPageInSection
            return borders.EnableFirstPageInSection
        else:
            # 其他页：检查EnableOtherPagesInSection
            return borders.EnableOtherPagesInSection
    
    except Exception:
        # 如果无法访问属性，默认认为启用
        return True


def get_section_border_info(section):
    """
    获取section的边框详细信息
    
    Args:
        section: Section对象
    
    Returns:
        dict: {
            'shadow': bool,
            'enable_first_page': bool,
            'enable_other_pages': bool
        }
    """
    try:
        borders = section.Borders
        return {
            'shadow': borders.Shadow,
            'enable_first_page': borders.EnableFirstPageInSection,
            'enable_other_pages': borders.EnableOtherPagesInSection
        }
    except Exception:
        return {
            'shadow': False,
            'enable_first_page': False,
            'enable_other_pages': False
        }
