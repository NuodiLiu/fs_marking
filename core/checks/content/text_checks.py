"""
Text occurrence and position checks for content validation.
"""

from ..base_check import BaseCheck
from ...utils.utils import fuzzy_match


class TextOccurrenceCheck(BaseCheck):
    """
    Check how many times a specific text pattern appears in the document.
    
    Can filter by heading_only or body_only to restrict the search scope.
    """
    
    def __init__(self, text, threshold=0.7, heading_only=False, body_only=False, expected_count=1):
        """
        Args:
            text: Text pattern to search for
            threshold: Fuzzy match threshold (default 0.7)
            heading_only: Only check heading paragraphs (default False)
            body_only: Only check body paragraphs (default False)
            expected_count: Expected number of occurrences (default 1)
        """
        self.text = text
        self.threshold = threshold
        self.heading_only = heading_only
        self.body_only = body_only
        self.expected_count = expected_count
    
    def check(self, doc):
        """
        Count occurrences of text in document.
        
        Args:
            doc: Word document object
        
        Returns:
            tuple: (bool, str) - (passes, message)
        """
        
        count = 0
        match_length = int(len(self.text) * 1.5)
        
        for para in doc.Paragraphs:
            # Apply filters
            if self.heading_only and para.OutlineLevel == 10:  # 10 = body text
                continue
            if self.body_only and para.OutlineLevel != 10:  # not body text
                continue
            
            # Get paragraph text and check for match
            para_text = para.Range.Text.strip()
            if not para_text:
                continue
            
            # Check if text matches at the start of paragraph
            compare_text = para_text[:match_length]
            if fuzzy_match(self.text, compare_text, self.threshold):
                count += 1
        
        passes = (count == self.expected_count)
        
        filter_desc = ""
        if self.heading_only:
            filter_desc = " in headings"
        elif self.body_only:
            filter_desc = " in body paragraphs"
        
        message = f"Text '{self.text[:30]}...' appears {count} time(s){filter_desc} (expected {self.expected_count})"
        
        return passes, message


class LastHeadingCheck(BaseCheck):
    """
    Check if the last heading in the document matches expected text.
    """
    
    def __init__(self, text, threshold=0.7):
        """
        Args:
            text: Expected heading text
            threshold: Fuzzy match threshold (default 0.7)
        """
        self.text = text
        self.threshold = threshold
    
    def check(self, doc):
        """
        Verify the last heading matches expected text.
        
        Args:
            doc: Word document object
        
        Returns:
            tuple: (bool, str) - (passes, message)
        """
        
        # Find the last heading by iterating backwards
        last_heading = None
        for i in range(doc.Paragraphs.Count, 0, -1):
            para = doc.Paragraphs(i)
            if para.OutlineLevel != 10:  # Not body text = heading
                last_heading = para
                break
        
        if not last_heading:
            return False, "No headings found in document"
        
        # Check if last heading matches expected text
        heading_text = last_heading.Range.Text.strip()
        match_length = int(len(self.text) * 1.5)
        compare_text = heading_text[:match_length]
        
        passes = fuzzy_match(self.text, compare_text, self.threshold)
        
        if passes:
            message = f"Last heading matches '{self.text}'"
        else:
            message = f"Last heading is '{heading_text[:50]}' (expected '{self.text}')"
        
        return passes, message


class LastParagraphCheck(BaseCheck):
    """
    Check if the last body paragraph matches expected text.
    """
    
    def __init__(self, text, threshold=0.7):
        """
        Args:
            text: Expected paragraph start text
            threshold: Fuzzy match threshold (default 0.7)
        """
        self.text = text
        self.threshold = threshold
    
    def check(self, doc):
        """
        Verify the last body paragraph matches expected text.
        
        Args:
            doc: Word document object
        
        Returns:
            tuple: (bool, str) - (passes, message)
        """
        
        # Find the last body paragraph by iterating backwards
        last_body_para = None
        for i in range(doc.Paragraphs.Count, 0, -1):
            para = doc.Paragraphs(i)
            if para.OutlineLevel == 10:  # Body text
                para_text = para.Range.Text.strip()
                if para_text:  # Non-empty
                    last_body_para = para
                    break
        
        if not last_body_para:
            return False, "No body paragraphs found in document"
        
        # Check if last body paragraph matches expected text
        para_text = last_body_para.Range.Text.strip()
        match_length = int(len(self.text) * 1.5)
        compare_text = para_text[:match_length]
        
        passes = fuzzy_match(self.text, compare_text, self.threshold)
        
        if passes:
            message = f"Last body paragraph matches '{self.text[:30]}...'"
        else:
            message = f"Last body paragraph is '{para_text[:50]}...' (expected '{self.text[:30]}...')"
        
        return passes, message
