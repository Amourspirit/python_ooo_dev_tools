from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XTextInputStream2

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.io.text_input_stream_partial import TextInputStreamPartial
from ooodev.adapter.io.active_data_sink_partial import ActiveDataSinkPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface

# see tests.test_adapter.test_ucb.test_simple_file_access.test_simple_file_access


class TextInputStream2Partial(TextInputStreamPartial, ActiveDataSinkPartial):
    """
    Partial Class XTextInputStream2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextInputStream2, interface: UnoInterface | None = XTextInputStream2) -> None:
        """
        Constructor

        Args:
            component (XTextInputStream2): UNO Component that implements ``com.sun.star.io.XTextInputStream2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextInputStream2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        TextInputStreamPartial.__init__(self, component=component, interface=None)
        ActiveDataSinkPartial.__init__(self, component=component, interface=None)
