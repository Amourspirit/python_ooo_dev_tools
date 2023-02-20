from __future__ import annotations

"""
Use caution when using this Module.

It is critical that all UNO enums extended using ``DeletedEnumMeta``
have a static method ``_get_deleted_attribs()`` that returns a tuple of string.
If ``_get_deleted_attribs()`` is not present a recursion error is raised.

Usage Example.

.. code-block:: python

    class AnchorType(
            metaclass=DeletedEnumMeta,
            type_name="com.sun.star.text.TextContentAnchorType",
            name_space="com.sun.star.text",
        ):
            @staticmethod
            def _get_deleted_attribs() -> Tuple[str]:
                return ("AT_FRAME",)
"""
from typing import Any
import uno
from ooo.helper.enum_helper import UnoEnumMeta

from ..exceptions import ex as mEx


class DeletedUnoEnumMeta(UnoEnumMeta):
    def __getattr__(cls, __name: str) -> uno.Enum | Any:
        if __name in cls._get_deleted_attribs():
            cls_name = cls.__name__
            accessed_via = f"Enum {cls_name!r}"
            raise mEx.DeletedAttributeError(f"attribute {__name!r} of {accessed_via} has been deleted")

        value = super().__getattr__(__name)
        return value
