from typing import Callable, List, Any

class MultilevelCheckPoint:
    def __init__(self, description: str, func: Callable[[List[Any]], bool]):
        self.description = description
        self.func = func
        self.passed = False
        self.error = ""
        self.needs_review = False

    def run(self, paragraphs):
        try:
            self.passed = self.func(paragraphs)
            return self.passed
        except Exception as e:
            self.error = f"Exception: {str(e)}"
            self.needs_review = True
            return False
