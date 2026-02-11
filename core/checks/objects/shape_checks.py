# core/checks/objects/shape_checks.py

from core.checks.base_check import BaseCheck
from core.utils.utils import fuzzy_match
from win32com.client import constants


class PageHasShapeCheck(BaseCheck):
    """检查指定页面是否有shape（插入图形）"""
    
    def __init__(self, page_num=1):
        """
        Args:
            page_num: 页码（1-based）
        """
        self.page_num = page_num
    
    def check(self, doc):
        try:
            shapes_found = 0
            
            for shape in doc.Shapes:
                try:
                    shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                    if shape_page == self.page_num:
                        shapes_found += 1
                except Exception as e:
                    continue
            
            if shapes_found > 0:
                return (True, f"✅ Page {self.page_num} has {shapes_found} shape(s)")
            else:
                return (False, f"❌ Page {self.page_num} has no shapes")
        
        except Exception as e:
            return (False, f"❌ Error checking shapes: {str(e)}")


class ShapeContainsTextCheck(BaseCheck):
    """检查指定页面的shape是否包含指定文字（任何可以包含文字的插入图形）"""
    
    def __init__(self, expected_text="", fuzzy=True, threshold=0.8, page_num=None):
        """
        Args:
            expected_text: 期望的文字内容
            fuzzy: 是否使用模糊匹配
            threshold: 模糊匹配阈值（0-1）
            page_num: 可选，页码（1-based），如果为None则搜索整个文档
        """
        self.page_num = page_num
        self.expected_text = expected_text
        self.fuzzy = fuzzy
        self.threshold = threshold
    
    def check(self, doc):
        """检查页面上的插入图形是否包含指定文字"""
        try:
            errors = []
            shapes_on_page = 0
            shapes_with_text = 0
            
            for i, shape in enumerate(doc.Shapes):
                try:
                    # 检查页码（如果指定了）
                    if self.page_num is not None:
                        shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                        if shape_page != self.page_num:
                            continue
                    
                    shapes_on_page += 1
                    
                    # 尝试从 shape 中提取文字
                    actual_text = None
                    try:
                        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                            actual_text = shape.TextFrame.TextRange.Text.strip()
                    except:
                        pass
                    
                    if not actual_text:
                        continue
                    
                    shapes_with_text += 1
                    
                    # 检查文本匹配
                    if self.fuzzy:
                        if fuzzy_match(self.expected_text, actual_text, self.threshold):
                            return (True, f"✅ Found shape with matching text: '{actual_text[:50]}'")
                        else:
                            errors.append(f"Shape text '{actual_text[:50]}' does not match")
                    else:
                        if self.expected_text in actual_text:
                            return (True, f"✅ Found shape containing: '{self.expected_text}'")
                        else:
                            errors.append(f"Shape text '{actual_text[:50]}' does not contain target")
                        
                except Exception as e:
                    continue
            
            # 没有找到匹配的
            page_desc = f"on page {self.page_num}" if self.page_num else "in document"
            if shapes_on_page == 0:
                return (False, f"❌ No shapes found {page_desc}")
            elif shapes_with_text == 0:
                return (False, f"❌ No shapes with text found {page_desc} (found {shapes_on_page} shapes)")
            else:
                return (False, f"❌ No matching shape found (checked {shapes_with_text} shapes with text)")
        
        except Exception as e:
            return (False, f"❌ Error checking shape text: {str(e)}")


class ShapeCountCheck(BaseCheck):
    """检查指定页面的shape数量（插入图形）"""
    
    def __init__(self, page_num=1, expected_count=1, mode="exact"):
        """
        Args:
            page_num: 页码（1-based）
            expected_count: 期望的shape数量
            mode: "exact"（精确）, "min"（至少）, "max"（最多）
        """
        self.page_num = page_num
        self.expected_count = expected_count
        self.mode = mode
    
    def check(self, doc):
        try:
            count = 0
            for shape in doc.Shapes:
                try:
                    shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                    if shape_page == self.page_num:
                        count += 1
                except:
                    continue
            
            if self.mode == "exact":
                if count == self.expected_count:
                    return (True, f"✅ Page {self.page_num} has exactly {count} shape(s)")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} shape(s), expected {self.expected_count}")
            
            elif self.mode == "min":
                if count >= self.expected_count:
                    return (True, f"✅ Page {self.page_num} has {count} shape(s) (>= {self.expected_count})")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} shape(s), expected at least {self.expected_count}")
            
            elif self.mode == "max":
                if count <= self.expected_count:
                    return (True, f"✅ Page {self.page_num} has {count} shape(s) (<= {self.expected_count})")
                else:
                    return (False, f"❌ Page {self.page_num} has {count} shape(s), expected at most {self.expected_count}")
            
            return (False, f"❌ Invalid mode: {self.mode}")
        
        except Exception as e:
            return (False, f"❌ Error checking shape count: {str(e)}")


class ShapeTypeCheck(BaseCheck):
    """检查shape类型（textbox, picture, wordart等）"""
    
    def __init__(self, shape_type="textbox", text_contains=None, fuzzy=True, threshold=0.7, page_num=None):
        """
        Args:
            shape_type: 期望类型 "textbox", "picture", "wordart", "autoshape"
            text_contains: 可选，用于识别特定shape（如果页面有多个）
            fuzzy: 文本匹配是否模糊
            threshold: 模糊匹配阈值
            page_num: 可选，页码（1-based），如果为None则搜索整个文档
        """
        self.page_num = page_num
        self.shape_type = shape_type.lower()
        self.text_contains = text_contains
        self.fuzzy = fuzzy
        self.threshold = threshold
    
    def check(self, doc):
        try:
            # Shape type constants
            msoTextBox = 17
            msoPicture = 13
            msoAutoShape = 1
            
            type_map = {
                "textbox": msoTextBox,
                "picture": msoPicture,
                "autoshape": msoAutoShape
            }
            
            expected_type = type_map.get(self.shape_type)
            if expected_type is None:
                return (False, f"❌ Unknown shape type: {self.shape_type}")
            
            shapes_on_page = []
            for shape in doc.Shapes:
                try:
                    if self.page_num is not None:
                        shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                        if shape_page == self.page_num:
                            shapes_on_page.append(shape)
                    else:
                        shapes_on_page.append(shape)
                except:
                    continue
            
            if not shapes_on_page:
                page_desc = f"on page {self.page_num}" if self.page_num else "in document"
                return (False, f"❌ No shapes found {page_desc}")
            
            # 如果指定了文本，先找到包含该文本的shape
            target_shapes = shapes_on_page
            if self.text_contains:
                target_shapes = []
                for shape in shapes_on_page:
                    try:
                        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                            actual_text = shape.TextFrame.TextRange.Text.strip()
                            if self.fuzzy:
                                if fuzzy_match(self.text_contains, actual_text, self.threshold):
                                    target_shapes.append(shape)
                            else:
                                if self.text_contains in actual_text:
                                    target_shapes.append(shape)
                    except:
                        continue
                
                if not target_shapes:
                    return (False, f"❌ No shapes found containing text '{self.text_contains[:30]}...'")
            
            # 检查shape类型
            for shape in target_shapes:
                try:
                    if shape.Type == expected_type:
                        page_desc = f"on page {self.page_num}" if self.page_num else "in document"
                        return (True, f"✅ Found {self.shape_type} {page_desc}")
                except Exception as e:
                    continue
            
            # 没有找到匹配类型
            type_names = []
            for shape in target_shapes:
                try:
                    type_names.append(f"type={shape.Type}")
                except:
                    pass
            
            return (False, f"❌ Found shapes but none are {self.shape_type} ({', '.join(type_names[:3])})")
        
        except Exception as e:
            return (False, f"❌ Error checking shape type: {str(e)}")


class ShapePositionCheck(BaseCheck):
    """检查shape位置是否在指定标题下方"""
    
    def __init__(self, after_heading="", heading_threshold=0.7, 
                 text_contains=None, text_threshold=0.7, tolerance_pt=100, page_num=None):
        """
        Args:
            after_heading: 标题文本
            heading_threshold: 标题匹配阈值
            text_contains: 用于识别特定shape的文本
            text_threshold: shape文本匹配阈值
            tolerance_pt: 允许的距离误差（磅）
            page_num: 可选，页码，如果为None则在找到标题的页面上搜索shape
        """
        self.page_num = page_num
        self.after_heading = after_heading
        self.heading_threshold = heading_threshold
        self.text_contains = text_contains
        self.text_threshold = text_threshold
        self.tolerance_pt = tolerance_pt
    
    def check(self, doc):
        try:
            # 1. 找到目标标题段落
            heading_para = None
            heading_page = None
            match_length = int(len(self.after_heading) * 1.5)
            
            for para in doc.Paragraphs:
                if para.OutlineLevel != 10:  # 是标题
                    para_text = para.Range.Text.strip()
                    compare_text = para_text[:match_length]
                    if fuzzy_match(self.after_heading, compare_text, self.heading_threshold):
                        para_page = para.Range.Information(constants.wdActiveEndPageNumber)
                        # 如果指定了页码，检查是否匹配；否则接受任何页面
                        if self.page_num is None or para_page == self.page_num:
                            heading_para = para
                            heading_page = para_page
                            break
            
            if not heading_para:
                page_desc = f"on page {self.page_num}" if self.page_num else "in document"
                return (False, f"❌ Heading '{self.after_heading}' not found {page_desc}")
            
            # 2. 找到目标shape（在标题所在页面或指定页面）
            target_shape = None
            search_page = self.page_num if self.page_num is not None else heading_page
            
            for shape in doc.Shapes:
                try:
                    shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                    if shape_page != search_page:
                        continue
                    
                    if self.text_contains:
                        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                            actual_text = shape.TextFrame.TextRange.Text.strip()
                            if fuzzy_match(self.text_contains, actual_text, self.text_threshold):
                                target_shape = shape
                                break
                    else:
                        target_shape = shape
                        break
                except:
                    continue
            
            if not target_shape:
                if self.text_contains:
                    return (False, f"❌ No shape found containing '{self.text_contains[:30]}...' on page {search_page}")
                else:
                    return (False, f"❌ No shape found on page {search_page}")
            
            # 3. 检查shape位置：比较anchor位置和垂直坐标
            # Method 1: Check if shape anchor comes after heading in document flow
            heading_range_end = heading_para.Range.End
            shape_anchor_start = target_shape.Anchor.Start
            
            if shape_anchor_start <= heading_range_end:
                return (False, f"❌ Shape anchor is not after heading (shape={shape_anchor_start}, heading_end={heading_range_end})")
            
            # Method 2: Check vertical position
            heading_top = heading_para.Range.Information(constants.wdVerticalPositionRelativeToPage)
            shape_top = target_shape.Top
            
            if shape_top <= heading_top:
                return (False, f"❌ Shape (top={shape_top:.1f}pt) is not below heading (top={heading_top:.1f}pt)")
            
            # Both checks passed
            return (True, f"✅ Shape is positioned below heading '{self.after_heading}'")
        
        except Exception as e:
            return (False, f"❌ Error checking shape position: {str(e)}")


class ShapeWrapFormatCheck(BaseCheck):
    """检查shape的文字环绕格式"""
    
    def __init__(self, wrap_type="top_bottom", text_contains=None, 
                 fuzzy=True, threshold=0.7, page_num=None):
        """
        Args:
            wrap_type: 环绕类型 "top_bottom", "square", "tight", "through", "inline"
            text_contains: 用于识别特定shape的文本
            fuzzy: 文本匹配是否模糊
            threshold: 阈值
            page_num: 可选，页码，如果为None则搜索整个文档
        """
        self.page_num = page_num
        self.wrap_type = wrap_type.lower()
        self.text_contains = text_contains
        self.fuzzy = fuzzy
        self.threshold = threshold
    
    def check(self, doc):
        try:
            # Wrap format constants
            wdWrapTopBottom = 1
            wdWrapSquare = 2
            wdWrapTight = 3
            wdWrapThrough = 4
            wdWrapInline = 7
            
            wrap_map = {
                "top_bottom": wdWrapTopBottom,
                "square": wdWrapSquare,
                "tight": wdWrapTight,
                "through": wdWrapThrough,
                "inline": wdWrapInline
            }
            
            expected_wrap = wrap_map.get(self.wrap_type)
            if expected_wrap is None:
                return (False, f"❌ Unknown wrap type: {self.wrap_type}")
            
            # 找到目标shape
            target_shape = None
            shapes_on_page = []
            
            for shape in doc.Shapes:
                try:
                    if self.page_num is not None:
                        shape_page = shape.Anchor.Information(constants.wdActiveEndPageNumber)
                        if shape_page != self.page_num:
                            continue
                    
                    shapes_on_page.append(shape)
                    
                    if self.text_contains:
                        if hasattr(shape, 'TextFrame') and shape.TextFrame.HasText:
                            actual_text = shape.TextFrame.TextRange.Text.strip()
                            if self.fuzzy:
                                if fuzzy_match(self.text_contains, actual_text, self.threshold):
                                    target_shape = shape
                                    break
                            else:
                                if self.text_contains in actual_text:
                                    target_shape = shape
                                    break
                    else:
                        target_shape = shape
                        break
                except:
                    continue
            
            if not shapes_on_page:
                page_desc = f"on page {self.page_num}" if self.page_num else "in document"
                return (False, f"❌ No shapes found {page_desc}")
            
            if not target_shape:
                if self.text_contains:
                    return (False, f"❌ No shape found containing '{self.text_contains[:30]}...'")
                else:
                    return (False, f"❌ No shape found")
            
            # 检查环绕格式
            try:
                actual_wrap = target_shape.WrapFormat.Type
                if actual_wrap == expected_wrap:
                    wrap_name = self.wrap_type.replace("_", " ").title()
                    return (True, f"✅ Shape has {wrap_name} text wrapping")
                else:
                    # 映射实际值到名称
                    reverse_map = {v: k for k, v in wrap_map.items()}
                    actual_name = reverse_map.get(actual_wrap, f"type={actual_wrap}")
                    expected_name = self.wrap_type
                    return (False, f"❌ Shape has {actual_name} wrapping, expected {expected_name}")
            except Exception as e:
                return (False, f"❌ Cannot read wrap format: {str(e)}")
        
        except Exception as e:
            return (False, f"❌ Error checking shape wrap format: {str(e)}")
