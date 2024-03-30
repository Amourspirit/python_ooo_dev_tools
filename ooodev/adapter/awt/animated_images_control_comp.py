from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.animation_partial import AnimationPartial

if TYPE_CHECKING:
    from com.sun.star.awt import AnimatedImagesControl


class AnimatedImagesControlComp(UnoControlComp, AnimationPartial):

    def __init__(self, component: AnimatedImagesControl):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.AnimatedImagesControl`` service.
        """
        UnoControlComp.__init__(self, component=component)
        AnimationPartial.__init__(self, component=self.component, interface=None)

    @property
    def component(self) -> AnimatedImagesControl:
        """AnimatedImagesControl Component"""
        # pylint: disable=no-member
        return cast("AnimatedImagesControl", self._ComponentBase__get_component())  # type: ignore
