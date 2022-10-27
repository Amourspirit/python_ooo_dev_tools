from abc import abstractmethod
from enum import Enum


class KindBase(str, Enum):
    def __str__(self) -> str:
        return self.value
    
    @abstractmethod
    def to_namespace(self) -> str:
        ...