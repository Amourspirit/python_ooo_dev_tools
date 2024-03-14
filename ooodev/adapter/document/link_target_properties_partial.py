from __future__ import annotations
from typing import TYPE_CHECKING
import uno


if TYPE_CHECKING:
    from com.sun.star.document import LinkTarget


class LinkTargetPropertiesPartial:
    """
    Partial class for LinkTarget.

    See Also:
        `API LinkTarget <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1document_1_1LinkTarget.html>`_
    """

    def __init__(self, component: LinkTarget) -> None:
        """
        Constructor

        Args:
            component (LinkTarget): UNO Component that implements ``com.sun.star.drawing.LinkTarget`` service.
        """
        self.__component = component

    # region LinkTarget
    @property
    def link_display_name(self) -> str:
        """
        Gets a human readable name for this object that could be displayed in a user interface.
        """
        return self.__component.LinkDisplayName

    # endregion LinkTarget
