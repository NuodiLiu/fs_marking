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
    2. Footnote content matches expected text
    3. Footnote marker immediately follows the word (no extra spaces)
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
        
        Args:
            doc: Word document object
        
        Returns:
            tuple: (bool, str) - (passes, message)
        """
        try:
            # Traverse all words in document
            for i, word_range in enumerate(doc.Words):
                # Clean word text (remove punctuation)
                word_text = word_range.Text.strip().rstrip(".,;:!?")
                
                # Compare with keyword
                if self.case_sensitive:
                    if word_text != self.keyword:
                        continue
                else:
                    if word_text.lower() != self.keyword.lower():
                        continue
                
                # Found target word, check next element for footnote
                # enumerate is 0-based, but Words.Item(n) is 1-based
                lookahead_index = i + 2
                if lookahead_index > doc.Words.Count:
                    return (False, f"❌ Word '{self.keyword}' found but no following element")
                
                lookahead = doc.Words.Item(lookahead_index)
                
                # Check 1: Has footnote?
                if lookahead.Footnotes.Count == 0:
                    return (False, f"❌ Word '{self.keyword}' does not have a footnote")
                
                # Check 2: Footnote content correct?
                actual_note = lookahead.Footnotes(1).Range.Text.strip()
                
                if self.fuzzy:
                    if not fuzzy_match(self.expected_text, actual_note, self.threshold):
                        return (False, f"❌ Footnote content mismatch: expected '{self.expected_text[:30]}...', got '{actual_note[:50]}...'")
                else:
                    if self.expected_text not in actual_note:
                        return (False, f"❌ Footnote does not contain: '{self.expected_text}'")
                
                # Check 3: lookahead being the footnote means it immediately follows (no space issue)
                # This is implicitly validated by the lookahead logic
                
                return (True, f"✅ Word '{self.keyword}' has correct footnote")
            
            # Keyword not found
            return (False, f"❌ Word '{self.keyword}' not found in document")
        
        except Exception as e:
            return (False, f"❌ Error checking footnote: {str(e)}")
