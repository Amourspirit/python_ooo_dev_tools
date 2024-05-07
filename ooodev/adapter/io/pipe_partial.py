from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XPipe

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.io.output_stream_partial import OutputStreamPartial
from ooodev.adapter.io.input_stream_partial import InputStreamPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PipePartial(OutputStreamPartial, InputStreamPartial):
    """
    Partial Class XPipe.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XPipe, interface: UnoInterface | None = XPipe) -> None:
        """
        Constructor

        Args:
            component (XPipe): UNO Component that implements ``com.sun.star.io.XPipe`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPipe``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        OutputStreamPartial.__init__(self, component=component, interface=None)
        InputStreamPartial.__init__(self, component=component, interface=None)

    # region XPipe

    # endregion XPipe
