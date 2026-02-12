# core/writers/feedback_writer.py

import os

class FeedbackWriter:
    def __init__(self, base_dir="logs"):
        self.base_dir = base_dir

    def write(self, zid: str, result: dict):
        student_dir = os.path.join(self.base_dir, zid)
        os.makedirs(student_dir, exist_ok=True)

        feedback_path = os.path.join(student_dir, "feedback.txt")
        with open(feedback_path, "w", encoding="utf-8") as f:
            f.write(f"📄 Feedback for {zid}\n")
            f.write(f"Total Score: {result['total']}\n\n")
            if result.get("needs_review", False):
                f.write("\n⚠️ Manual review suggested.\n")
                
            for rule in result["results"]:
                # Skip placeholder rules
                if rule.get("placeholder", False):
                    continue
                
                mark = rule["mark"]
                name = rule["name"]
                errors = rule.get("errors", [])
                needs_review = rule.get("needs_review", False)
                if mark == 0:
                    f.write(f"❌ {name}: 0 mark\n")
                    for err in errors:
                        f.write(f"    - {err}\n")
                elif needs_review:
                    f.write(f"⚠️ {name}: {mark} mark (needs manual review)\n")
                else:
                    f.write(f"✔️ {name}: {mark} mark\n")
            
