from __future__ import annotations

# https://tinyurl.com/2eftghrw
class static_prop:
    def __init__(self, getter):
        self.__getter = getter

    def __get__(self, obj, objtype):
        return self.__getter(objtype)

    @staticmethod
    def __call__(getter_fn):
        return static_prop(getter_fn)
