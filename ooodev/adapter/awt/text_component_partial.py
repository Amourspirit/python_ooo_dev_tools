from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XTextComponent
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XTextListener
    from com.sun.star.awt import Selection
    from ooodev.utils.type_var import UnoInterface


class TextComponentPartial:
    """
    Partial class for XTextComponent.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextComponent, interface: UnoInterface | None = XTextComponent) -> None:
        """
        Constructor

        Args:
            component (XTextComponent): UNO Component that implements ``com.sun.star.awt.XTextComponent`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextComponent``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTextComponent
    def add_text_listener(self, listener: XTextListener) -> None:
        """
        Registers a text event listener.
        """
        self.__component.addTextListener(listener)

    def get_max_text_len(self) -> int:
        """
        Gets the currently set maximum text length.
        """
        return self.__component.getMaxTextLen()

    def get_selected_text(self) -> str:
        """
        Gets the currently selected text.
        """
        return self.__component.getSelectedText()

    def get_selection(self) -> Selection:
        """
        Gets the current user selection.
        """
        return self.__component.getSelection()

    def get_text(self) -> str:
        """
        Gets the text of the component.
        """
        return self.__component.getText()

    def insert_text(self, sel: Selection, text: str) -> None:
        """
        Inserts text at the specified position.

        Returns:
            None:

        Hint:
            ``Selection`` can be imported from ``ooo.dyn.awt.selection``.
        """
        self.__component.insertText(sel, text)

    def is_editable(self) -> bool:
        """
        Gets if the text is editable by the user.
        """
        return self.__component.isEditable()

    def remove_text_listener(self, listener: XTextListener) -> None:
        """
        Un-registers a text event listener.
        """
        self.__component.removeTextListener(listener)

    def set_editable(self, editable: bool) -> None:
        """
        Makes the text editable for the user or read-only.
        """
        self.__component.setEditable(editable)

    def set_max_text_len(self, length: int) -> None:
        """
        Sets the maximum text length.
        """
        self.__component.setMaxTextLen(length)

    def set_selection(self, selection: Selection) -> None:
        """
        Sets the user selection.

        Returns:
            None:

        Hint:
            ``Selection`` can be imported from ``ooo.dyn.awt.selection``.
        """
        self.__component.setSelection(selection)

    def set_text(self, text: str) -> None:
        """
        Sets the text of the component.
        """
        self.__component.setText(text)

    # endregion XTextComponent
