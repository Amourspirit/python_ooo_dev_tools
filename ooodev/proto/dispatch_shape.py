from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from typing_extensions import Protocol
    from com.sun.star.drawing import XShape
    from com.sun.star.drawing import XDrawPage

    class DispatchShape(Protocol):
        """This is strictly for GUI helpers such as ``ooo-dev-tools-gui-win`` that have a create_dispatch_shape method"""

        def __call__(self, slide: XDrawPage, shape_dispatch: str) -> XShape | None: ...

else:
    DispatchShape = object
