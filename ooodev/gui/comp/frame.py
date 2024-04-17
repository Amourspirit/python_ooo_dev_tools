from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.frame import frame2_partial
from ooodev.adapter.frame.frame_action_events import FrameActionEvents
from ooodev.gui.comp.layout_manager import LayoutManager
from ooodev.utils import info as mInfo
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.loader.inst.lo_inst import LoInst


if TYPE_CHECKING:
    from com.sun.star.frame import Frame as UnoFrame  # service
    from com.sun.star.frame import XFrame
    from com.sun.star.uno import XInterface


class _Frame(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.frame.Frame",)

    # region XFramesSupplier Overrides
    def get_active_frame(self) -> Frame | None:
        """
        Gets the current active frame of this container (not of any other available supplier)

        This may be the frame itself. The active frame is defined as the frame which contains (recursively) the window with the focus.
        If no window within the frame contains the focus, this method returns the last frame which had the focus.
        If no containing window ever had the focus, the first frame within this frame is returned.
        """
        frm = self.__component.getActiveFrame()
        if frm is None:
            return None
        return Frame(component=frm, lo_inst=self.lo_inst)  # type: ignore

    def set_active_frame(self, frame: XFrame | Frame) -> None:
        """
        Is called on activation of a direct sub-frame.

        This method is only allowed to be called by a sub-frame according to ``XFrame.activate()`` or ``XFramesSupplier.setActiveFrame()``.
        After this call ``XFramesSupplier.getActiveFrame()`` will return the frame specified by Frame.

        In general this method first calls the method ``XFramesSupplier.setActiveFrame()`` at the creator frame with this as the current argument.
        Then it broadcasts the FrameActionEvent ``FrameAction.FRAME_ACTIVATED``.

        Args:
            frame (XFrame | FrameComp): The frame to set as active.

        Note:
            Given parameter Frame must already exist inside the container (e.g., inserted by using ``XFrames.append()``)
        """
        if mInfo.Info.is_instance(frame, Frame):
            self.__component.setActiveFrame(frame.component)
        else:
            self.__component.setActiveFrame(frame)

    # endregion XFramesSupplier Overrides

    # region Frame2Partial Overrides
    @property
    def layout_manager(self) -> LayoutManager:
        """
        Provides access to the LayoutManager of the frame.

        When setting can be instance of ``XLayoutManager2`` or ``LayoutManagerComp``.
        """

        lm = self.component.LayoutManager
        if lm is None:
            return None  # type: ignore
        return LayoutManager(component=lm, lo_inst=self.lo_inst)  # type: ignore

    @layout_manager.setter
    def layout_manager(self, value: XInterface | ComponentProp) -> None:
        if mInfo.Info.is_instance(value, ComponentProp):
            self.component.LayoutManager = value.component
        else:
            self.component.LayoutManager = value  # type: ignore

    # endregion Frame2Partial Overrides


class Frame(
    _Frame,
    frame2_partial.Frame2Partial,
    FrameActionEvents,
    CompDefaultsPartial,
    LoInstPropsPartial,
    # child_partial.ChildPartial,
):
    """
    Class for managing Frame Component.

    Note:
        This is a Dynamic class that is created at runtime.
        This means that the class is created at runtime and not defined in the source code.
        In addition, the class may be created with additional classes implemented.

        The Type hints for this class at design time may not be accurate.
        To check if a class implements a specific interface, use the ``isinstance`` function
        or :py:meth:`~.InterfacePartial.is_supported_interface` methods which is always available in this class.
    """

    # pylint: disable=unused-argument

    def __new__(cls, component: Any, **kwargs):
        lo_inst = kwargs.get("lo_inst", None)
        builder = get_builder(component=component)
        if lo_inst is not None:
            builder.lo_inst = lo_inst
        builder_helper.builder_add_comp_defaults(builder)
        builder_helper.builder_add_property_change_implement(builder)
        builder_helper.builder_add_property_veto_implement(builder)
        builder_helper.builder_add_lo_inst_props_partial(builder)

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.gui.comp.frame.Frame",
            base_class=_Frame,
        )
        return inst

    def __init__(self, component: Any, *, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that supports ``com.sun.star.frame.Frame`` service.
            lo_inst (LoInst, optional): Instance of the Lo. Defaults to ``None``.
        """
        # this it not actually called as __new__ is overridden
        pass

    # region Properties

    @property
    def component(self) -> UnoFrame:
        """Frame Component"""
        # pylint: disable=no-member
        return cast("UnoFrame", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component=component)
    builder.auto_interface()
    # builder = frame2_partial.get_builder(component=component)

    builder.add_event(
        module_name="ooodev.adapter.frame.frame_action_events",
        class_name="FrameActionEvents",
        uno_name="com.sun.star.frame.XFrame",
        optional=True,
    )
    builder.add_event(
        module_name="ooodev.adapter.frame.title_change_events",
        class_name="TitleChangeEvents",
        uno_name="com.sun.star.frame.XTitleChangeBroadcaster",
        optional=True,
    )
    # ooodev.adapter.frame.title_change_events

    return builder
