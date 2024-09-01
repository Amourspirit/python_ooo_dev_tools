"""
Use caution when using this Module.

It is critical that all UNO enums extended using ``DeletedEnumMeta``
have a static method ``_get_deleted_attribs()`` that returns a tuple of string.
If ``_get_deleted_attribs()`` is not present a recursion error is raised.

Usage Example.

.. code-block:: python

    class AnchorKind(
            metaclass=DeletedEnumMeta,
            type_name="com.sun.star.text.TextContentAnchorType",
            name_space="com.sun.star.text",
        ):
            @staticmethod
            def _get_deleted_attribs() -> Tuple[str]:
                return ("AT_FRAME",)
"""

from __future__ import annotations
from typing import Any
import uno
from ooo.helper.enum_helper import UnoEnumMeta
from ooo.helper.enum_helper import ConstEnumMeta

from ooodev.exceptions import ex as mEx


# since oooenv version 3.0.0 the UnoEnumMeta has been updated to use __getattribute__ instead of __getattr__
# This implementation is also backward compatible with the previous implementation.
# However, the previous implementation should not be used needed going forward.


class DeletedUnoEnumMeta(UnoEnumMeta):
    """Descriptor to raise an exception when an UNO Enum attribute is accessed after deletion."""

    # def __getattr__(cls, name: str) -> uno.Enum | Any:
    #     if name in cls._get_deleted_attribs():  # type: ignore
    #         cls_name = cls.__name__
    #         accessed_via = f"Enum {cls_name!r}"
    #         raise mEx.DeletedAttributeError(f"attribute {name!r} of {accessed_via} has been deleted")

    #     return super().__getattr__(name)  # type: ignore

    def __getattribute__(cls, name: str) -> Any:
        # object.__getattribute__ must be used here or a recursion error will be raised.
        restricted = object.__getattribute__(cls, "_get_deleted_attribs")()
        if name in restricted:  # type: ignore
            cls_name = cls.__name__
            accessed_via = f"Enum {cls_name!r}"
            raise mEx.DeletedAttributeError(f"attribute {name!r} of {accessed_via} has been deleted")
        return super().__getattribute__(name)


# The ConstEnumMeta can use the __getattr__ to check for deleted attributes


class DeletedUnoConstEnumMeta(ConstEnumMeta):
    """Descriptor to raise an exception when an attribute UNO Const is accessed after deletion."""

    def __getattr__(cls, name: str) -> uno.Enum | Any:
        if name in cls._get_deleted_attribs():  # type: ignore
            cls_name = cls.__name__
            accessed_via = f"Enum {cls_name!r}"
            raise mEx.DeletedAttributeError(f"attribute {name!r} of {accessed_via} has been deleted")

        return super().__getattr__(name)  # type: ignore
