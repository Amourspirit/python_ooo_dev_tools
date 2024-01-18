from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.drawing import XGraphicExportFilter

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

from ooodev.adapter.document.exporter_partial import ExporterPartial
from ooodev.adapter.document.filter_partial import FilterPartial
from ooodev.adapter.document.mime_type_info_partial import MimeTypeInfoPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class GraphicExportFilterPartial(ExporterPartial, FilterPartial, MimeTypeInfoPartial):
    """
    Partial class for XGraphicExportFilter.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XGraphicExportFilter, interface: UnoInterface | None = XGraphicExportFilter) -> None:
        """
        Constructor

        Args:
            component (XGraphicExportFilter): UNO Component that implements ``com.sun.star.container.XGraphicExportFilter`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XGraphicExportFilter``.
        """

        self.__interface = interface
        self.__validate(component)
        ExporterPartial.__init__(self, component, interface)
        FilterPartial.__init__(self, component, interface)
        MimeTypeInfoPartial.__init__(self, component, interface)

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)
