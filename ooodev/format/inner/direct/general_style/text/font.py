from __future__ import annotations
from typing import Any
import contextlib
import uno


from com.sun.star.beans import XPropertySet
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.direct.write.char.font.font import Font as CharFont


class Font(CharFont):
    """
    General font class for text.

    No validation is done other then checking of the object supports the ``XPropertySet`` interface.

    If validation fails then nothing happens.

    Any errors are suppressed.
    """

    def apply(self, obj: Any, **kwargs: Any) -> None:
        if obj is None:
            return
        if mLo.Lo.qi(XPropertySet, obj) is None:
            return
        validate = bool(kwargs.get("validate", False))
        super().apply(obj=obj, validate=validate, **kwargs)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        # set properties. Can be overridden in child classes
        # may be useful to wrap in try statements in child classes
        with contextlib.suppress(Exception):
            mProps.Props.set(obj, **kwargs)
