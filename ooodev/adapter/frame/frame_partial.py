from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XFrame

from ooodev.adapter.lang.component_partial import ComponentPartial

if TYPE_CHECKING:
    from ooodev.utils.builder.default_builder import DefaultBuilder
    from com.sun.star.frame import XFrameActionListener
    from com.sun.star.awt import XWindow
    from com.sun.star.frame import XController
    from com.sun.star.frame import XFramesSupplier
    from ooodev.utils.type_var import UnoInterface


class FramePartial(ComponentPartial):
    """
    Partial class for XFrame.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFrame, interface: UnoInterface | None = XFrame) -> None:
        """
        Constructor

        Args:
            component (XFrame ): UNO Component that implements ``com.sun.star.frame.XFrame`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFrame``.
        """

        ComponentPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XFrame
    def activate(self) -> None:
        """
        activates this frame and thus the component within.

        At first the frame sets itself as the active frame of its creator by calling ``XFramesSupplier.setActiveFrame()``,
        then it broadcasts a ``FrameActionEvent`` with ``FrameAction.FRAME_ACTIVATED``.

        The component within this frame may listen to this event to grab the focus on activation;
        for simple components this can be done by the FrameLoader.

        Finally, most frames may grab the focus to one of its windows or forward the activation to a sub-frame.
        """
        self.__component.activate()

    def add_frame_action_listener(self, listener: XFrameActionListener) -> None:
        """
        Registers an event listener, which will be called when certain things happen to the components within this frame or within sub-frames of this frame.

        E.g., it is possible to determine instantiation/destruction and activation/deactivation of components.
        """
        self.__component.addFrameActionListener(listener)

    def context_changed(self) -> None:
        """
        Notifies the frame that the context of the controller within this frame changed (i.e. the selection).

        According to a call to this interface, the frame calls ``XFrameActionListener.frameAction()`` with ``FrameAction.CONTEXT_CHANGED``
        to all listeners which are registered using ``XFrame.addFrameActionListener()``.
        For external controllers this event can be used to re-query dispatches.
        """
        self.__component.contextChanged()

    def deactivate(self) -> None:
        """
        Is called by the creator frame when another sub-frame gets activated.

        At first the frame deactivates its active sub-frame, if any. Then broadcasts a FrameActionEvent with FrameAction.FRAME_DEACTIVATING.
        """
        self.__component.deactivate()

    def find_frame(self, target_frame_name: str, search_flags: int) -> XFrame:
        """
        searches for a frame with the specified name.

        Frames may contain other frames (e.g., a frameset) and may be contained in other frames.
        This hierarchy is searched with this method.
        First some special names are taken into account, i.e. ``_self``, ``_top``, ``_blank`` etc.
        SearchFlags is ignored when comparing these names with TargetFrameName; further steps are controlled by SearchFlags.
        If allowed, the name of the frame itself is compared with the desired one, and then ( again if allowed ) the method is called for all children of the frame.
        Finally may be called for the siblings and then for parent frame (if allowed).

        List of special target names:

        If no frame with the given name is found, a new top frame is created; if this is allowed by a special flag ``FrameSearchFlag.CREATE``.
        The new frame also gets the desired name.

        See Also:
            `com.sun.star.frame.FrameSearchFlag <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1frame_1_1FrameSearchFlag.html>`__
        """
        return self.__component.findFrame(target_frame_name, search_flags)

    def get_component_window(self) -> XWindow:
        """
        Provides access to the component window

        Note:
            Don't dispose this window - the frame is the owner of it.
        """
        return self.__component.getComponentWindow()

    def get_container_window(self) -> XWindow:
        """
        Provides access to the container window of the frame.

        Normally this is used as the parent window of the component window.
        """
        return self.__component.getContainerWindow()

    def get_controller(self) -> XController:
        """
        Provides access to the controller

        Note:
            Don't dispose it - the frame is the owner of it. Use ``XController.getFrame()`` to dispose the frame after you the controller agreed with a ``XController.suspend()`` call.
        """
        return self.__component.getController()

    def get_creator(self) -> XFramesSupplier:
        """
        Provides access to the creator (parent) of this frame.
        """
        return self.__component.getCreator()

    def get_name(self) -> str:
        """
        Get the name property of this frame.
        """
        return self.__component.getName()

    def initialize(self, window: XWindow) -> None:
        """
        Is called to initialize the frame within a window - the container window.

        This window will be used as parent for the component window and to support some UI relevant features of the frame service.
        Note: Re-parenting mustn't supported by a real frame implementation! It's designed for initializing - not for setting.

        This frame will take over ownership of the window referred from xWindow. Thus, the previous owner is not allowed to dispose this window anymore.
        """
        self.__component.initialize(window)

    def is_active(self) -> bool:
        """
        determines if the frame is active.
        """
        return self.__component.isActive()

    def is_top(self) -> bool:
        """
        determines if the frame is a top frame.

        In general a top frame is the frame which is a direct child of a task frame or which does not have a parent. Possible frame searches must stop the search at such a frame unless the flag FrameSearchFlag.TASKS is set.
        """
        return self.__component.isTop()

    def remove_frame_action_listener(self, listener: XFrameActionListener) -> None:
        """
        Un-registers an event listener
        """
        self.__component.removeFrameActionListener(listener)

    def set_component(self, window: XWindow, controller: XController) -> bool:
        """
        Sets a new component into the frame or release an existing one from a frame.

        A valid component window should be a child of the frame container window.

        Simple components may implement a ``com.sun.star.awt.XWindow`` only. In this case no controller must be given here.
        """
        return self.__component.setComponent(window, controller)

    def set_creator(self, creator: XFramesSupplier) -> None:
        """
        sets the frame container that created this frame.

        Only the creator is allowed to call this method. But creator doesn't mean the implementation which creates this instance ... it means the parent frame of the frame hierarchy. Because; normally a frame should be created by using the API and is necessary for searches inside the tree (e.g. XFrame.findFrame())
        """
        self.__component.setCreator(creator)

    def set_name(self, name: str) -> None:
        """
        Sets the name of the frame.

        Normally the name of the frame is set initially (e.g. by the creator).
        The name of a frame will be used for identifying it if a frame search was started.
        These searches can be forced by:

        Note:
            Special targets like ``_blank``, ``_self`` etc. are not allowed. That's why frame names shouldn't start with a sign ``_``.
        """
        self.__component.setName(name)

    # endregion XFrame


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.frame.XFrame", False)
    builder.set_omit("com.sun.star.lang.XComponent")
    return builder
