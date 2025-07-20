class StdoutWriter:
    def write(self, zid: str, result: dict):
        print(f"\n🧾 Debug Output for {zid}")
        print(f"Total Score: {result['total']}")

        for rule in result["results"]:
            name = rule["name"]
            mark = rule["mark"]

            # Example: color green for full marks, red for zero, default otherwise
            if mark == 0:
                color = "\033[91m"  # Red
            elif mark == rule.get("max_mark", mark):  # If you have max_mark info
                color = "\033[92m"  # Green
            else:
                color = "\033[0m"   # Default

            print(f"  • {name}: {color}{mark} mark\033[0m")

            errors = rule.get("errors", [])
            for err in errors:
                print(f"     - {err}")

        if result.get("needs_review", False):
            print("⚠️  Manual review required.")