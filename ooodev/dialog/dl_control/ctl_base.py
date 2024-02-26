# region imports
from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import
import unohelper

from com.sun.star.awt import XView
from com.sun.star.awt import XWindow
from com.sun.star.beans import XPropertySet
from ooo.dyn.awt.pos_size import PosSize

from ooodev.events.args.generic_args import GenericArgs
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
from ooodev.loader import lo as mLo
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.utils.partial.view_prop_partial import ViewPropPartial
from ooodev.adapter.awt.uno_control_dialog_element_partial import UnoControlDialogElementPartial


if TYPE_CHECKING:
    from com.sun.star.awt import XControlModel
    from com.sun.star.awt import XControl
    from com.sun.star.awt import UnoControlDialogElement  # service
    from com.sun.star.beans import XMultiPropertySet
    from ooodev.proto.style_obj import StyleT
# endregion imports

# pylint: disable=unused-argument


class CtlBase(unohelper.Base, LoInstPropsPartial, ViewPropPartial, ModelPropPartial, EventsPartial):
    """Control Base Class"""

    # region Dunder Methods
    def __init__(self, ctl: Any) -> None:
        unohelper.Base.__init__(self)
        LoInstPropsPartial.__init__(self, lo_inst=mLo.Lo.current_lo)
        EventsPartial.__init__(self)
        model = ctl.getModel()
        ViewPropPartial.__init__(self, obj=ctl)
        ModelPropPartial.__init__(self, obj=model)
        self._set_control(ctl)

    def _set_control(self, ctl: Any) -> None:
        self._ctl_view = ctl

    # endregion Dunder Methods

    # region other methods

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    def get_view(self) -> XControl:
        return self._ctl_view

    get_view_ctl = get_view

    def get_uno_srv_name(self) -> str:
        """Get Uno service name"""
        raise NotImplementedError

    def get_control_props(self) -> XPropertySet | None:
        """
        Gets property set for a control model

        Args:
            control_model (Any): control model

        Returns:
            XPropertySet | None: Property set
        """
        try:
            return mLo.Lo.qi(XPropertySet, self.get_model())
        except Exception:
            return None

    # endregion other methods

    # region Other Methods
    def get_model(self) -> XControlModel:
        """Gets the Model for the control"""
        return self.get_view().getModel()

    # endregion Other Methods

    # region Properties

    @property
    def visible(self) -> bool:
        """Gets/Sets the visible state for the control"""
        try:
            model = cast(Any, self.get_view().getModel())
            return model.EnableVisible
        except Exception:
            return True

    @visible.setter
    def visible(self, value: bool) -> None:
        with contextlib.suppress(Exception):
            model = cast(Any, self.get_model())
            model.EnableVisible = value

    @property
    def x(self) -> int:
        """Gets/Sets the x position for the control"""
        view = mLo.Lo.qi(XWindow, self.get_view(), True)
        return view.getPosSize().X

    @x.setter
    def x(self, value: int) -> None:
        win = mLo.Lo.qi(XWindow, self.get_view(), True)
        size_pos = win.getPosSize()
        win.setPosSize(value, size_pos.Y, size_pos.Width, size_pos.Height, PosSize.X)

    @property
    def y(self) -> int:
        """Gets/Sets the y position for the control"""
        view = mLo.Lo.qi(XWindow, self.get_view(), True)
        return view.getPosSize().Y

    @y.setter
    def y(self, value: int) -> None:
        view = mLo.Lo.qi(XWindow, self.get_view(), True)
        size = view.getPosSize()
        view.setPosSize(size.X, value, size.Width, size.Height, PosSize.Y)

    # endregion Properties


class CtlListenerBase(
    CtlBase,
    UnoControlDialogElementPartial,
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
    """Dialog Control Listener Base Class"""

    # region Dunder Methods
    def __init__(self, ctl: Any) -> None:
        CtlBase.__init__(self, ctl)
        UnoControlDialogElementPartial.__init__(self)
        generic_args = self._get_generic_args()
        # The Events callback methods are invoked when any event is added or removed.
        FocusEvents.__init__(self, trigger_args=generic_args, cb=self._on_focus_listener_add_remove)
        KeyEvents.__init__(self, trigger_args=generic_args, cb=self._on_key_events_listener_add_remove)
        MouseEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_listener_add_remove)
        MouseMotionEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_motion_listener_add_remove)
        PaintEvents.__init__(self, trigger_args=generic_args, cb=self._on_paint_listener_add_remove)
        WindowEvents.__init__(self, trigger_args=generic_args, cb=self._on_window_event_listener_add_remove)
        model = self.get_model()
        PropertyChangeImplement.__init__(self, component=cast(XPropertySet, model), trigger_args=generic_args)
        PropertiesChangeImplement.__init__(self, component=cast("XMultiPropertySet", model), trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=cast(XPropertySet, model), trigger_args=generic_args)

    # endregion Dunder Methods

    # region Overrides
    def get_uno_srv_name(self) -> str:
        """Get Uno service name"""
        raise NotImplementedError

    # endregion Overrides

    # region Lazy Listeners

    # Listeners such as mouse, mouse motion, focus, key, paint, and window are added lazily.
    # Some listeners such as mouse motion may be expensive.
    # Each time the mouse moves, the listener is invoked.
    # By lazy loading listeners are only added when needed.
    # For example:
    #     ctl_button_ok.add_event_mouse_entered(on_mouse_entered)
    # This would call _on_mouse_listener_add_remove() below, which would add the mouse listener to the class in a lazy manor.

    def _on_mouse_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addMouseListener(self.events_listener_mouse)
        event.remove_callback = True

    def _on_mouse_motion_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addMouseMotionListener(self.events_listener_mouse_motion)
        event.remove_callback = True

    def _on_focus_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addFocusListener(self.events_listener_focus)
        event.remove_callback = True

    def _on_key_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addKeyListener(self.events_listener_key)
        event.remove_callback = True

    def _on_paint_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addPaintListener(self.events_listener_paint)
        event.remove_callback = True

    def _on_window_event_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        view = cast(Any, self.get_view_ctl())
        view.addWindowListener(self.events_listener_window)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region other methods
    def get_property_set(self) -> XPropertySet:
        """Gets the property set for this control"""
        return mLo.Lo.qi(XPropertySet, self.get_model(), True)

    # endregion other methods

    # region UnoControlDialogElementPartial Overrides
    @property
    def width(self) -> int:
        """Gets the width of the control"""
        view = mLo.Lo.qi(XView, self.get_view(), True)
        return view.getSize().Width

    @width.setter
    def width(self, value: int) -> None:
        win = mLo.Lo.qi(XWindow, self.get_view(), True)
        pos_size = win.getPosSize()
        pos_size.Width = value
        win.setPosSize(pos_size.X, pos_size.Y, pos_size.Width, pos_size.Height, PosSize.WIDTH)

    @property
    def height(self) -> int:
        """Gets/Sets the height of the control"""
        view = mLo.Lo.qi(XView, self.get_view(), True)
        return view.getSize().Height

    @height.setter
    def height(self, value: int) -> None:
        win = mLo.Lo.qi(XWindow, self.get_view(), True)
        pos_size = win.getPosSize()
        pos_size.Height = value
        win.setPosSize(pos_size.X, pos_size.Y, pos_size.Width, pos_size.Height, PosSize.HEIGHT)

    # endregion UnoControlDialogElementPartial Overrides


class DialogControlBase(CtlListenerBase):
    """Dialog Control Base Class. Only for Controls that have a model that can be added to a dialog"""

    def __init__(self, ctl: Any) -> None:
        CtlListenerBase.__init__(self, ctl)

    # region Overrides
    def get_uno_srv_name(self) -> str:
        """Get Uno service name"""
        raise NotImplementedError

    # endregion Overrides

    # region other methods
    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind"""
        return DialogControlKind.from_value(self.get_uno_srv_name())

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind"""
        with contextlib.suppress(Exception):
            kind = self.get_control_kind()
            return DialogControlNamedKind.from_str(kind.name)
        return DialogControlNamedKind.UNKNOWN

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

    # region Properties
    @property
    def tab_index(self) -> int:
        """Gets/Sets the tab index"""
        with contextlib.suppress(Exception):
            model = cast("UnoControlDialogElement", self.get_model())
            return model.TabIndex
        return -1

    @tab_index.setter
    def tab_index(self, value: int) -> None:
        with contextlib.suppress(Exception):
            model = cast("UnoControlDialogElement", self.get_model())
            model.TabIndex = value

    @property
    def step(self) -> int:
        """Gets/Sets the step"""
        with contextlib.suppress(Exception):
            model = cast("UnoControlDialogElement", self.get_model())
            return model.Step
        return 0

    @step.setter
    def step(self, value: int) -> None:
        with contextlib.suppress(Exception):
            model = cast("UnoControlDialogElement", self.get_model())
            model.Step = value

    @property
    def tag(self) -> str:
        """Gets/Sets the tag"""
        with contextlib.suppress(Exception):
            model = cast("UnoControlDialogElement", self.get_model())
            return model.Tag
        return ""

    @tag.setter
    def tag(self, value: str) -> None:
        with contextlib.suppress(Exception):
            model = cast("UnoControlDialogElement", self.get_model())
            model.Tag = value

    @property
    def tip_text(self) -> str:
        """Gets/Sets the tip text"""
        with contextlib.suppress(Exception):
            model = cast(Any, self.get_model())
            return model.HelpText
        return ""

    @tip_text.setter
    def tip_text(self, value: str) -> None:
        with contextlib.suppress(Exception):
            model = cast(Any, self.get_model())
            model.HelpText = value

    # useful alias
    help_text = tip_text
    # endregion Properties
