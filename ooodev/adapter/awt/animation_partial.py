from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XAnimation

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class AnimationPartial:
    """
    Partial class for XAnimation.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XAnimation, interface: UnoInterface | None = XAnimation) -> None:
        """
        Constructor

        Args:
            component (XAnimation): UNO Component that implements ``com.sun.star.awt.XAnimation`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XAnimation``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XAnimation
    def is_animation_running(self) -> bool:
        """
        Gets whether the animation is currently running
        """
        return self.__component.isAnimationRunning()

    def start_animation(self) -> None:
        """
        starts the animation
        """
        self.__component.startAnimation()

    def stop_animation(self) -> None:
        """
        Stops the animation
        """
        self.__component.stopAnimation()

    # endregion XAnimation
