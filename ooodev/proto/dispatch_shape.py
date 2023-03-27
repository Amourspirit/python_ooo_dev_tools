from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object

import uno
from com.sun.star.drawing import XShape
from com.sun.star.drawing import XDrawPage

from ..utils.data_type.window_title import WindowTitle


class DispatchShape(Protocol):
    """This is strictly for GUI helpers such as ``ooo-dev-tools-gui-win`` that have a create_dispatch_shape method"""

    def __call__(self, slide: XDrawPage, shape_dispatch: str) -> XShape | None:
        ...
