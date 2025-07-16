# core/rules/strict_heading1_rule.py
from core.rules.heading_rules.heading_text_base import StrictHeadingRuleBase

class StrictHeading1Rule(StrictHeadingRuleBase):
    def __init__(self, mark=1):
        expected = [
            "BIOLOGY AND NATURAL HISTORY",
            "URBAN ADAPTATION AND PUBLIC PERCEPTION",
            "CULTURAL SIGNIFICANCE, CONSERVATION, AND FUTURE PROSPECTS",
            "MAIN POINTS",
            "SUMMARY",
            "REFERENCES"
        ]
        super().__init__("Strict Heading 1 Check", mark, expected, style_name="Heading 1")
