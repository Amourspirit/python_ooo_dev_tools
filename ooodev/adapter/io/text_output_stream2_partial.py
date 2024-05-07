from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XTextOutputStream2

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.io.text_output_stream_partial import TextOutputStreamPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TextOutputStream2Partial(TextOutputStreamPartial):
    """
    Partial Class XTextOutputStream2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextOutputStream2, interface: UnoInterface | None = XTextOutputStream2) -> None:
        """
        Constructor

        Args:
            component (XTextOutputStream2): UNO Component that implements ``com.sun.star.io.XTextOutputStream2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextOutputStream2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        TextOutputStreamPartial.__init__(self, component=component, interface=None)
