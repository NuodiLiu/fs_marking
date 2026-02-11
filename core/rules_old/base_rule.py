from abc import ABC, abstractmethod
from core.logging.logger import setup_logger

class BaseRule(ABC):
    def __init__(self, name: str = None, mark: int = 1):
        self.name = name or self.__class__.__name__
        self.mark = mark
        self.logger = setup_logger(self.name)

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