from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno
from com.sun.star.beans import XPropertySet

from ooo.dyn.awt.pos_size import PosSize

# from ooodev.adapter.lang.event_events import EventEvents
from ooodev.adapter.adapter_base import GenericArgs
from ooodev.adapter.awt.focus_events import FocusEvents
from ooodev.adapter.awt.key_events import KeyEvents
from ooodev.adapter.awt.mouse_events import MouseEvents
from ooodev.adapter.awt.mouse_motion_events import MouseMotionEvents
from ooodev.adapter.awt.paint_events import PaintEvents
from ooodev.adapter.awt.window_events import WindowEvents
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import lo as mLo
from ooodev.utils.kind.form_component_kind import FormComponentKind

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.awt import UnoControlModel  # service
    from com.sun.star.awt import UnoControl  # service
    from ooodev.proto.style_obj import StyleT


class FormCtlBase(
    FocusEvents,
    KeyEvents,
    MouseEvents,
    MouseMotionEvents,
    PaintEvents,
    WindowEvents,
    PropertyChangeImplement,
    PropertiesChangeImplement,
    VetoableChangeImplement,
):
    """Base class for all form controls"""

    # region init
    def __init__(self, ctl: XControl) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlModel): Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        self._set_control(ctl)
        model = self.get_model()
        generic_args = self._get_generic_args()
        FocusEvents.__init__(self, trigger_args=generic_args, cb=self._on_focus_listener_add_remove)
        KeyEvents.__init__(self, trigger_args=generic_args, cb=self._on_key_events_listener_add_remove)
        MouseEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_listener_add_remove)
        MouseMotionEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_motion_listener_add_remove)
        PaintEvents.__init__(self, trigger_args=generic_args, cb=self._on_paint_listener_add_remove)
        WindowEvents.__init__(self, trigger_args=generic_args, cb=self._on_window_event_listener_add_remove)
        PropertyChangeImplement.__init__(self, component=model)
        PropertiesChangeImplement.__init__(self, component=model)
        VetoableChangeImplement.__init__(self, component=model)

    # endregion init

    # region Lazy Listeners
    # def _on_event_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
    #     # will only ever fire once
    #     model = self.get_model()
    #     model.addEventListener(self.events_listener_event)
    #     event.remove_callback = True

    def _on_mouse_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addMouseListener(self.events_listener_mouse)
        event.remove_callback = True

    def _on_mouse_motion_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addMouseMotionListener(self.events_listener_mouse_motion)
        event.remove_callback = True

    def _on_focus_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addFocusListener(self.events_listener_focus)
        event.remove_callback = True

    def _on_key_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addKeyListener(self.events_listener_key)
        event.remove_callback = True

    def _on_paint_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addPaintListener(self.events_listener_paint)
        event.remove_callback = True

    def _on_window_event_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = self.get_view()
        view.addWindowListener(self.events_listener_window)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region other methods
    def get_id(self) -> int:
        """
        Gets class id for this control.

        Returns:
            int: Class Id if found, Otherwise ``-1``
        """
        props = self.get_property_set()
        with contextlib.suppress(Exception):
            return int(props.getPropertyValue("ClassId"))
        mLo.Lo.print("No class ID found for form component")
        return -1

    def get_model(self) -> UnoControlModel:
        """Gets the model for this control"""
        return cast("UnoControlModel", self.get_control().getModel())

    def get_view(self) -> UnoControl:
        """Gets the view of this control"""
        return cast("UnoControl", self.get_control().getView())

    def get_control(self) -> XControl:
        """Gets the control"""
        return self.__control

    def _set_control(self, ctl: XControl) -> None:
        self.__control = ctl

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    def get_property_set(self) -> XPropertySet:
        """Gets the property set for this control"""
        return mLo.Lo.qi(XPropertySet, self.get_model(), True)

    def apply_styles(self, *styles: StyleT) -> None:
        """
        Applies styles to control

        Args:
            *styles: Styles to apply
        """
        model = self.get_model()
        for style in styles:
            style.apply(model)

    # endregion other methods

    # region Overrides
    def get_uno_srv_name(self) -> str:
        """Get Uno service name"""
        return self.get_form_component_kind().to_namespace()

    # endregion Overrides

    # region Abstract Methods
    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        raise NotImplementedError("Must override get_form_component_kind")

    # endregion Abstract Methods

    # region Properties

    @property
    def width(self) -> int:
        """Gets the width of the control"""
        return self.get_view().getSize().Width

    @width.setter
    def width(self, value: int) -> None:
        view = self.get_view()
        pos_size = view.getPosSize()
        pos_size.Width = value
        view.setPosSize(pos_size.X, pos_size.Y, pos_size.Width, pos_size.Height, PosSize.WIDTH)

    @property
    def height(self) -> int:
        """Gets/Sets the height of the control"""
        return self.get_view().getSize().Height

    @height.setter
    def height(self, value: int) -> None:
        view = self.get_view()
        pos_size = view.getPosSize()
        pos_size.Height = value
        view.setPosSize(pos_size.X, pos_size.Y, pos_size.Width, pos_size.Height, PosSize.HEIGHT)

    @property
    def x(self) -> int:
        """Gets/Sets the x position for the control"""
        return self.get_view().getPosSize().X

    @x.setter
    def x(self, value: int) -> None:
        view = self.get_view()
        pos_size = view.getPosSize()
        pos_size.X = value
        view.setPosSize(pos_size.X, pos_size.Y, pos_size.Width, pos_size.Height, PosSize.X)

    @property
    def y(self) -> int:
        """Gets/Sets the y position for the control"""
        return self.get_view().getPosSize().Y

    @y.setter
    def y(self, value: int) -> None:
        view = self.get_view()
        pos_size = view.getPosSize()
        pos_size.Y = value
        view.setPosSize(pos_size.X, pos_size.Y, pos_size.Width, pos_size.Height, PosSize.Y)

    @property
    def name(self) -> str:
        """Gets the name for the control model"""
        return self.get_model().Name

    @property
    def tab_index(self) -> int:
        """Gets/Sets the tab index"""
        return self.get_model().TabIndex

    @tab_index.setter
    def tab_index(self, value: int) -> None:
        self.get_model().TabIndex = value

    @property
    def tag(self) -> str:
        """Gets/Sets the tag"""
        return self.get_model().Tag

    @tag.setter
    def tag(self, value: str) -> None:
        self.get_model().Tag = value

    # endregion Properties
