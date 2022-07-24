# coding: utf-8
from typing import Any, Callable, TypeVar, Generic

# import functools

T = TypeVar("T")


class classproperty(Generic[T]):
    def __init__(self, fget: Callable[[Any], T], fset=None) -> None:
        # https://stackoverflow.com/questions/71018821/property-decorator-type-hints
        self.fget = fget
        self.fset = fset
        # https://stackoverflow.com/a/17705456/1171746
        # no need to update wrapper in this case.
        # becasue this is a decorater in a metaclass
        # i am not sure yet if the doc string is accessible at all
        # functools.update_wrapper(self, fget)  ## TA-DA! ##

    def __get__(self, obj, objtype=None) -> T:
        val = self.fget(obj)
        return val

    def __set__(self, obj, value: T) -> T:
        if not self.fset:
            raise AttributeError("can't set attribute")
        type_ = type(obj)
        return self.fset.__get__(obj, type_)(value)

    def setter(self, func):
        self.fset = func
        return self


class classinstanceproperty(Generic[T]):
    def __init__(self, fget: Callable[[Any], T], fset=None) -> None:
        self.fget = fget
        self.fset = fset
        # functools.update_wrapper(self, fget)  ## TA-DA! ##

    def __get__(self, obj, objtype=None) -> T:
        return self.fget(obj)

    def __set__(self, obj, value: T) -> T:
        if not self.fset:
            raise AttributeError("can't set attribute")
        type_ = type(obj)
        return self.fset.__get__(obj, type_)(value)

    def setter(self, func):
        self.fset = func
        return self


class StaticProperty(type):
    def __new__(self, name, bases, props):
        class_properties = {}
        to_remove = {}
        for key, value in props.items():
            if isinstance(value, (classproperty, classinstanceproperty)):
                class_properties[key] = value
                if isinstance(value, classproperty):
                    to_remove[key] = value

        for key in to_remove:
            props.pop(key)

        HoistMeta = type("HoistMeta", (type,), class_properties)
        return HoistMeta(name, bases, props)
