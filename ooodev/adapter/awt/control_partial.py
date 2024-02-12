from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.awt import XControl

# should inherit from ComponentPartial, however, a circular import is occurring when
# docs are being built. So, we implement XEventListener methods here.
# from ooodev.adapter.lang.component_partial import ComponentPartial

if TYPE_CHECKING:
    from com.sun.star.lang import XEventListener
    from com.sun.star.awt import XControlModel
    from com.sun.star.awt import XToolkit
    from com.sun.star.awt import XView
    from com.sun.star.awt import XWindowPeer
    from com.sun.star.uno import XInterface
    from ooodev.utils.type_var import UnoInterface


class ControlPartial:
    """
    Partial Class for XControl.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XControl, interface: UnoInterface | None = XControl) -> None:
        """
        Constructor

        Args:
            component (XControl): UNO Component that implements ``com.sun.star.awt.XControl`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XControl``.
        """
        self.__component = component

    # region XControl
    def create_peer(self, toolkit: XToolkit, parent: XWindowPeer) -> None:
        """
        Creates a child window on the screen.

        If the parent is ``None``, then the desktop window of the toolkit is the parent.
        """
        self.__component.createPeer(toolkit, parent)

    def get_context(self) -> XInterface:
        """
        Gets the context of the control.
        """
        return self.__component.getContext()

    def get_model(self) -> XControlModel:
        """
        Gets the model for this control.
        """
        return self.__component.getModel()

    def get_peer(self) -> XWindowPeer:
        """
        Gets the peer which was previously created or set.
        """
        return self.__component.getPeer()

    def get_view(self) -> XView:
        """
        Gets the view of this control.
        """
        return self.__component.getView()

    def is_design_mode(self) -> bool:
        """
        Returns ``True`` if the control is in design mode, ``False`` otherwise.
        """
        return self.__component.isDesignMode()

    def is_transparent(self) -> bool:
        """
        Returns ``True`` if the control is transparent, ``False`` otherwise.
        """
        return self.__component.isTransparent()

    def set_context(self, ctx: XInterface) -> None:
        """
        Sets the context of the control.
        """
        self.__component.setContext(ctx)

    def set_design_mode(self, on: bool) -> None:
        """
        Sets the design mode for use in a design editor.

        Normally the control will be painted directly without a peer.
        """
        self.__component.setDesignMode(on)

    def set_model(self, model: XControlModel) -> bool:
        """
        Sets a model for the control.
        """
        return self.__component.setModel(model)

    # endregion XControl

    # region XComponent
    def add_event_listener(self, listener: XEventListener) -> None:
        """
        Adds an event listener to the component.

        Args:
            listener (XEventListener): The event listener to be added.
        """
        self.__component.addEventListener(listener)

    def remove_event_listener(self, listener: XEventListener) -> None:
        """
        Removes an event listener from the component.

        Args:
            listener (XEventListener): The event listener to be removed.
        """
        self.__component.removeEventListener(listener)

    def dispose(self) -> None:
        """
        Disposes the component.
        """
        self.__component.dispose()

    # endregion XComponent
