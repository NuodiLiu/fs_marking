import re
import difflib

def cm_to_points(cm: float) -> float:
    return cm * 28.35  # Word: 1cm = 28.35 pt

def validate_zid(zid):
    # Z1234567 or z1234567 or 1234567
    return re.fullmatch(r"(z\d{7}|\d{7})", zid, re.IGNORECASE)

def validate_config(config):
    if not hasattr(config, "RULES"):
        raise AttributeError("❌ The config object must define a RULES attribute.")

    if not isinstance(config.RULES, list):
        raise TypeError("❌ The RULES attribute must be a list.")

    for rule in config.RULES:
        if not hasattr(rule, "run") or not callable(rule.run):
            raise TypeError(f"❌ Each rule must implement a callable .run(doc) method. Invalid: {type(rule)}")

    print("✅ Config validation passed.")

def fuzzy_match(expected: str, actual: str, threshold: float = 0.8) -> bool:
    """
    Returns True if the actual string matches the expected string 
    with at least `threshold` similarity using SequenceMatcher.

    Similarity ranges from 0.0 (completely different) to 1.0 (identical).
    This function is case-insensitive.

    Examples (threshold = 0.8):

        expected = "The Australian Ibis"

        actual                          similarity   match?   notes
        --------------------------------------------------------------------
        "The Australian Ibis"              1.00       ✅     exact match
        "the australian ibis"              1.00       ✅     case-insensitive
        "The Australien Ibis"              0.94       ✅     minor typo
        "The AustraIian Ibis"              0.90       ✅     'l' mistaken for 'I'
        "The Austrailian Ibus"             0.86       ✅     2 typos, still acceptable
        "The Australib Ibis"               0.83       ✅     single letter misplaced
        "Australian Ibis"                  0.83       ✅     missing "The"
        "The Australia Ibis"               0.83       ✅     missing a character
        "The Ibiss"                        0.61       ❌     missing key words
        "Australian Bird"                  0.51       ❌     too different
        "Completely Different"             0.25       ❌     unrelated

    """
    ratio = difflib.SequenceMatcher(None, expected.lower(), actual.lower()).ratio()
    return ratio >= threshold
