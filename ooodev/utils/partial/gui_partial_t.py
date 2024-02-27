from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.frame import XController
    from com.sun.star.frame import XDispatchProviderInterception
    from com.sun.star.frame import XFrame
    from com.sun.star.view import XControlAccess
    from com.sun.star.view import XSelectionSupplier

    from typing_extensions import Protocol
else:
    Protocol = object


class GuiPartialT(Protocol):
    def get_current_controller(self) -> XController:
        """
        Gets controller from document.

        Returns:
            XController: controller.
        """
        ...

    def get_frame(self) -> XFrame:
        """
        Gets frame from doc.

        Returns:
            XFrame: document frame.
        """
        ...

    def get_control_access(self) -> XControlAccess:
        """
        Get control access from office document.

        Returns:
            XControlAccess: control access.
        """
        ...

    def get_selection_supplier(self) -> XSelectionSupplier:
        """
        Gets selection supplier

        Returns:
            XSelectionSupplier: Selection supplier
        """
        ...

    def get_dpi(self) -> XDispatchProviderInterception:
        """
        Gets Dispatch provider interception.

        Returns:
            XDispatchProviderInterception: Dispatch provider interception.
        """
        ...

    def activate(self) -> None:
        """
        Activates document window.
        """

        ...

    def maximize(self) -> None:
        """
        Maximizes document window.
        """
        ...

    def minimize(self) -> None:
        """
        Minimizes document window.
        """
        ...

    def add_item_to_toolbar(self, toolbar_name: str, item_name: str, im_fnm: str) -> None:
        """
        Add a user-defined icon and command to the start of the specified toolbar.

        Args:
            toolbar_name (str): toolbar name.
            item_name (str): item name.
            im_fnm (str): image file path.
        """
        ...

    def show_menu_bar(self) -> None:
        """
        Shows the main menu bar.
        """
        ...

    def hide_menu_bar(self) -> None:
        """
        Hides the main menu bar.
        """
        ...

    def hide_all_menu_bars(self) -> None:
        """
        Make all the toolbars invisible.
        """
        ...

    def toggle_menu_bar(self) -> None:
        """
        Toggles the main menu bar.
        """
        ...
