# core/writers/stdout_writer.py

class StdoutWriter:
    def write(self, zid: str, result: dict):
        print(f"\n🧾 Debug Output for {zid}")
        print(f"Total Score: {result['total']}")

        for rule in result["results"]:
            name = rule["name"]
            mark = rule["mark"]
            print(f"  • {name}: {mark} mark")

            errors = rule.get("errors", [])
            for err in errors:
                print(f"     - {err}")

            if result.get("needs_review", False):
                print("⚠️  Manual review required.")
