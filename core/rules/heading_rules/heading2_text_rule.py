# core/rules/strict_heading2_rule.py
from core.rules.heading_rules.heading_text_base import StrictHeadingRuleBase

class StrictHeading2Rule(StrictHeadingRuleBase):
    def __init__(self, mark=1):
        expected = [
            "Physical Characteristics",
            "Diet and Foraging Behaviour",
            "Breeding and Life Cycle",
            "From Wetlands to Concrete Jungles",
            "Survivors of the Anthropocene",
            "Public Image and Media Representation",
            "Cultural and Indigenous Significance",
            "Conservation Status and Concerns",
            "Coexistence and the Future"
        ]
        super().__init__("Strict Heading 2 Check", mark, expected, style_name="Heading 2")
