from abc import ABC, abstractmethod

class BaseRule(ABC):
    def __init__(self, name: str = None, mark: int = 1):
        self.name = name or self.__class__.__name__
        self.mark = mark

    @abstractmethod
    def run(self, doc) -> dict:
        """
        after run should return:
        {
            "name": str,
            "mark": int,
            "errors": List[str],
            "needs_review": bool
        }
        """
        pass