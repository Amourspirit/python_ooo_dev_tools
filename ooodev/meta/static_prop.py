from __future__ import annotations
from collections.abc import Callable
from typing import Generic, TypeVar


R = TypeVar("R")

# https://tinyurl.com/2eftghrw
# https://tinyurl.com/2o9y7bdl
class static_prop(Generic[R]):
    def __init__(self, getter: Callable[[], R]) -> None:
        self.__getter = getter

    def __get__(self, obj: object, objtype: type) -> R:
        return self.__getter()

    @staticmethod
    def __call__(getter_fn: Callable[[], R]) -> static_prop[R]:
        return static_prop(getter_fn)
