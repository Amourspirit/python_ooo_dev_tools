from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XFramesSupplier

from ooodev.adapter.frame.frame_partial import FramePartial

if TYPE_CHECKING:
    from com.sun.star.frame import XFrame
    from com.sun.star.frame import XFrames
    from ooodev.utils.type_var import UnoInterface


class FramesSupplierPartial(FramePartial):
    """
    Partial class for XFramesSupplier.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XFramesSupplier, interface: UnoInterface | None = XFramesSupplier) -> None:
        """
        Constructor

        Args:
            component (XFramesSupplier ): UNO Component that implements ``com.sun.star.frame.XFramesSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XFramesSupplier``.
        """

        FramePartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XFramesSupplier
    def get_active_frame(self) -> XFrame:
        """
        Gets the current active frame of this container (not of any other available supplier)

        This may be the frame itself. The active frame is defined as the frame which contains (recursively) the window with the focus.
        If no window within the frame contains the focus, this method returns the last frame which had the focus.
        If no containing window ever had the focus, the first frame within this frame is returned.
        """
        return self.__component.getActiveFrame()

    def get_frames(self) -> XFrames:
        """
        Provides access to this container and to all other ``XFramesSupplier`` which are available from this node of frame tree.
        """
        return self.__component.getFrames()

    def set_active_frame(self, frame: XFrame) -> None:
        """
        Is called on activation of a direct sub-frame.

        This method is only allowed to be called by a sub-frame according to ``XFrame.activate()`` or ``XFramesSupplier.setActiveFrame()``.
        After this call ``XFramesSupplier.getActiveFrame()`` will return the frame specified by Frame.

        In general this method first calls the method ``XFramesSupplier.setActiveFrame()`` at the creator frame with this as the current argument.
        Then it broadcasts the FrameActionEvent ``FrameAction.FRAME_ACTIVATED``.

        Note:
            Given parameter Frame must already exist inside the container (e.g., inserted by using ``XFrames.append()``)
        """
        self.__component.setActiveFrame(frame)

    # endregion XFramesSupplier
