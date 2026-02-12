"""
Footnote validation checks.
"""

from ..base_check import BaseCheck
from ...utils.utils import fuzzy_match


class FootnoteCheck(BaseCheck):
    """
    Check if a specific word has a footnote with correct content.
    
    Validates:
    1. Word has a footnote
    2. Footnote reference immediately follows the word (nothing in between)
    3. Footnote content matches expected text
    """
    
    def __init__(self, keyword="", expected_text="", fuzzy=True, threshold=0.75, case_sensitive=False):
        """
        Args:
            keyword: Word to search for (e.g., "effect")
            expected_text: Expected footnote content
            fuzzy: Use fuzzy matching for footnote content
            threshold: Fuzzy match threshold (0-1)
            case_sensitive: Case-sensitive word matching
        """
        self.keyword = keyword
        self.expected_text = expected_text
        self.fuzzy = fuzzy
        self.threshold = threshold
        self.case_sensitive = case_sensitive
    
    def check(self, doc):
        """
        Check if keyword has footnote with correct content.
        
        Uses a reverse approach: iterates through document footnotes and checks
        if the text immediately before each footnote reference is the keyword.
        This ensures nothing exists between the keyword and the footnote.
        
        Args:
            doc: Word document object
        
        Returns:
            tuple: (bool, str) - (passes, message)
        """
        try:
            # Check if document has any footnotes
            if doc.Footnotes.Count == 0:
                # Check if keyword exists at all
                if self._keyword_exists(doc):
                    return (False, f"❌ Word '{self.keyword}' found but document has no footnotes")
                return (False, f"❌ Word '{self.keyword}' not found in document")
            
            # Iterate through all footnotes and check what precedes each reference
            for fn in doc.Footnotes:
                ref_range = fn.Reference
                ref_start = ref_range.Start
                
                # Get the text immediately before the footnote reference
                keyword_len = len(self.keyword)
                if ref_start < keyword_len:
                    continue
                
                before_range = doc.Range(ref_start - keyword_len, ref_start)
                before_text = before_range.Text
                
                # Check if text before footnote matches the keyword
                if self.case_sensitive:
                    match = before_text == self.keyword
                else:
                    match = before_text.lower() == self.keyword.lower()
                
                if not match:
                    continue
                
                # Found keyword immediately before footnote reference — 
                # by construction, nothing is between them.
                
                # Validate footnote content
                actual_note = fn.Range.Text.strip()
                
                if self.fuzzy:
                    if not fuzzy_match(self.expected_text, actual_note, self.threshold):
                        return (False, 
                            f"❌ Footnote content mismatch: expected '{self.expected_text[:30]}...', "
                            f"got '{actual_note[:50]}...'")
                else:
                    if self.expected_text not in actual_note:
                        return (False, f"❌ Footnote does not contain: '{self.expected_text}'")
                
                return (True, f"✅ Word '{self.keyword}' has correct footnote")
            
            # No footnote found on the keyword — check if keyword exists
            if self._keyword_exists(doc):
                return (False, f"❌ Word '{self.keyword}' does not have a footnote immediately following it")
            
            return (False, f"❌ Word '{self.keyword}' not found in document")
        
        except Exception as e:
            return (False, f"❌ Error checking footnote: {str(e)}")
    
    def _keyword_exists(self, doc):
        """Check if the keyword exists anywhere in the document using Find."""
        try:
            find_range = doc.Range()
            find_range.Find.Text = self.keyword
            find_range.Find.MatchCase = self.case_sensitive
            find_range.Find.MatchWholeWord = True
            return find_range.Find.Execute()
        except Exception:
            return False
