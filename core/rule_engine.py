# core/rule_engine.py

def evaluate(doc, rules):
    """
    Execute all scoring rules and aggregate the results.

    Parameters:
        doc: Word document object (win32com)
        rules: A list of rule objects, each implementing .run(doc)

    Returns:
        dict: {
            "total": total score,
            "results": [
                {"name": rule name, "mark": score, "errors": [error messages], "needs_review": True}
            ]
        }
    """
    total = 0
    results = []

    for rule in rules:
        try:
            result = rule.run(doc)

            # make sure format is consitant
            result.setdefault("name", rule.__class__.__name__)
            result.setdefault("mark", 0)
            result.setdefault("errors", [])
            result.setdefault("needs_review", False)

            results.append(result)
            total += result["mark"]

        except Exception as e:
            # If a rule fails, it doesn't affect the rest
            results.append({
                "name": rule.__class__.__name__,
                "mark": 0,
                "errors": [f"Rule execution error: {str(e)}"]
            })

    return {
        "total": total,
        "results": results
    }