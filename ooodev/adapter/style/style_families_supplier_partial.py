from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.style import XStyleFamiliesSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from com.sun.star.container import XNameAccess


class StyleFamiliesSupplierPartial:
    """
    Partial class for XStyleFamiliesSupplier.
    """

    def __init__(
        self, component: XStyleFamiliesSupplier, interface: UnoInterface | None = XStyleFamiliesSupplier
    ) -> None:
        """
        Constructor

        Args:
            component (XStyleFamiliesSupplier): UNO Component that implements ``com.sun.star.style.XStyleFamiliesSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStyleFamiliesSupplier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XStyleFamiliesSupplier
    def get_style_families(self) -> XNameAccess:
        """
        This method returns the collection of style families available in the container document.
        """
        return self.__component.getStyleFamilies()

    # endregion XStyleFamiliesSupplier
