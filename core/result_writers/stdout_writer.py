# core/writers/stdout_writer.py

class StdoutWriter:
    def write(self, zid: str, result: dict):
        print(f"🧾 Debug Output for {zid}")
        print(f"Total Score: {result['total']}")

        for rule in result["results"]:
            # Skip placeholder rules
            if rule.get("placeholder", False):
                continue
            
            name = rule["name"]
            mark = rule["mark"]
            print(f"  • {name}: {mark} mark")

            errors = rule.get("errors", [])
            for err in errors:
                print(f"     - {err}")

        if result.get("needs_review", False):
            print("⚠️  Manual review required.")

        print()