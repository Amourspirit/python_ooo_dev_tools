"""
Disable methods in child classes.

Usage:

.. code-block:: python

    class Foo:
        def hello(self) -> None:
            print("Hello")

        def msg(self, msg: str) -> None:
            print(msg)

        @staticmethod
        def foo_ness() -> None:
            print("Foo is me")


    class Bar(Foo):
        msg = DisabledMethod()
        foo_ness = DisabledMethod()
"""

from __future__ import annotations

# https://stackoverflow.com/questions/231839/python-inheritance-how-to-disable-a-function
from ooodev.exceptions import ex as mEx


class DisabledMethod:
    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, instance, owner):
        cls_name = owner.__name__
        accessed_via = f"type object {cls_name!r}" if instance is None else f"{cls_name!r} object"
        raise mEx.DisabledMethodError(f"method {self.name!r} of {accessed_via} has been disabled.")
