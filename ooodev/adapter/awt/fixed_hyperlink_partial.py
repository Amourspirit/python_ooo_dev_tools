from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XFixedHyperlink

from ooodev.utils.kind.align_kind import AlignKind
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XActionListener
    from ooodev.utils.type_var import UnoInterface


class FixedHyperlinkPartial:
    """
    Partial class for XFixedHyperlink.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFixedHyperlink, interface: UnoInterface | None = XFixedHyperlink) -> None:
        """
        Constructor

        Args:
            component (XFixedHyperlink): UNO Component that implements ``com.sun.star.awt.XFixedHyperlink`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFixedHyperlink``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XFixedHyperlink
    def add_action_listener(self, listener: XActionListener) -> None:
        """
        Registers an event handler for click action event.
        """
        self.__component.addActionListener(listener)

    def get_alignment(self) -> AlignKind:
        """
        Gets the alignment of the text in the control.

        Returns:
            AlignKind: The alignment of the text.

        Hint:
            - ``AlignKind`` can be imported from ``ooodev.utils.kind.align_kind``.
        """
        return AlignKind(self.__component.getAlignment())

    def get_text(self) -> str:
        """
        Gets the text of the control.
        """
        return self.__component.getText()

    def get_url(self) -> str:
        """
        Gets the url of the control.
        """
        return self.__component.getURL()

    def remove_action_listener(self, listener: XActionListener) -> None:
        """
        Un-registers an event handler for click action event.
        """
        self.__component.removeActionListener(listener)

    def set_alignment(self, align: int | AlignKind) -> None:
        """
        Sets the alignment of the text in the control.

        Args:
            align (int | AlignKind): The alignment of the text.

        Hint:
            - ``AlignKind`` can be imported from ``ooodev.utils.kind.align_kind``.
        """
        self.__component.setAlignment(int(align))

    def set_text(self, text: str) -> None:
        """
        sets the text of the control.
        """
        self.__component.setText(text)

    def set_url(self, url: str) -> None:
        """
        Sets the url of the control.
        """
        self.__component.setURL(url)

    # endregion XFixedHyperlink
