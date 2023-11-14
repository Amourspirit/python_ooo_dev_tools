# region imports
from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING
import uno
import unohelper

from com.sun.star.awt import XView
from com.sun.star.awt import XWindow
from com.sun.star.beans import XPropertySet

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.adapter.awt.focus_events import FocusEvents
from ooodev.adapter.awt.key_events import KeyEvents
from ooodev.adapter.awt.mouse_events import MouseEvents
from ooodev.adapter.awt.mouse_motion_events import MouseMotionEvents
from ooodev.adapter.awt.paint_events import PaintEvents
from ooodev.adapter.awt.window_events import WindowEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils import lo as mLo

from ooo.dyn.awt.pos_size import PosSize

if TYPE_CHECKING:
    from com.sun.star.awt import XControlModel
    from com.sun.star.awt import XControl

# endregion imports


class CtlBase(unohelper.Base):
    """Control Base Class"""

    # region Dunder Methods
    def __init__(self, ctl: Any) -> None:
        self._ctl_view = ctl
        self._set_listeners = set()

    def __getattr__(self, name: str) -> Any:
        # this is mostly for backwards compatibility
        if hasattr(self._ctl_view, name):
            return getattr(self._ctl_view, name)
        raise AttributeError(name)

    # endregion Dunder Methods

    # region other methods

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    def get_view_ctl(self) -> XControl:
        return self._ctl_view

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
            return mLo.Lo.qi(XPropertySet, self.get_view_ctl().getModel())
        except Exception:
            return None

    # endregion other methods

    # region Lazy Listeners

    # Listeners such as mouse, mouse motion, focus, key, paint, and window are added lazily.
    # Some listeners such as mouse motion may be expensive.
    # Each time the mouse moves, the listener is invoked.
    # By lazy loading listeners are only added when needed.
    # For example:
    #     ctl_button_ok.add_event_mouse_entered(on_mouse_entered)
    # This would call _on_mouse_listener_add_remove() below, which would add the mouse listener to the class in a lazy manor.

    def _has_listener(self, key: str) -> bool:
        """Gets if the listener key has been added"""
        if not key:
            raise ValueError("key cannot be empty")
        return key in self._set_listeners

    def _add_listener(self, key: str) -> None:
        """Adds a listener key to the set of listeners"""
        if not key:
            raise ValueError("key cannot be empty")
        self._set_listeners.add(key)

    # endregion Lazy Listeners

    # region Other Methods
    def get_model(self) -> XControlModel:
        """Gets the Model for the control"""
        return self.get_view_ctl().getModel()

    # endregion Other Methods

    # region Properties
    @property
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        model = cast(Any, self.get_view_ctl().getModel())
        return model.Enabled if hasattr(model, "Enabled") else True

    @enabled.setter
    def enabled(self, value: bool) -> None:
        model = cast(Any, self.get_view_ctl().getModel())
        if hasattr(model, "Enabled"):
            model.Enabled = value

    @property
    def visible(self) -> bool:
        """Gets/Sets the visible state for the control"""
        try:
            model = cast(Any, self.get_view_ctl().getModel())
            return model.EnableVisible
        except Exception:
            return True

    @property
    def width(self) -> int:
        """Gets the width of the control"""
        view = mLo.Lo.qi(XView, self.get_view_ctl(), True)
        return view.getSize().Width

    @width.setter
    def width(self, value: int) -> None:
        win = mLo.Lo.qi(XWindow, self.get_view_ctl(), True)
        pos_size = win.getPosSize()
        pos_size.Width = value
        win.setPosSize(pos_size.X, pos_size.Y, pos_size.Width, pos_size.Height, PosSize.WIDTH)

    @property
    def height(self) -> int:
        """Gets/Sets the height of the control"""
        view = mLo.Lo.qi(XView, self.get_view_ctl(), True)
        return view.getSize().Height

    @height.setter
    def height(self, value: int) -> None:
        win = mLo.Lo.qi(XWindow, self.get_view_ctl(), True)
        pos_size = win.getPosSize()
        pos_size.Height = value
        win.setPosSize(pos_size.X, pos_size.Y, pos_size.Width, pos_size.Height, PosSize.HEIGHT)

    @property
    def x(self) -> int:
        """Gets/Sets the x position for the control"""
        view = mLo.Lo.qi(XWindow, self.get_view_ctl(), True)
        return view.getPosSize().X

    @x.setter
    def x(self, value: int) -> None:
        win = mLo.Lo.qi(XWindow, self.get_view_ctl(), True)
        size_pos = win.getPosSize()
        win.setPosSize(value, size_pos.Y, size_pos.Width, size_pos.Height, PosSize.X)

    @property
    def y(self) -> int:
        """Gets/Sets the y position for the control"""
        view = mLo.Lo.qi(XWindow, self.get_view_ctl(), True)
        return view.getPosSize().Y

    @y.setter
    def y(self, value: int) -> None:
        view = mLo.Lo.qi(XWindow, self.get_view_ctl(), True)
        size = view.getPosSize()
        view.setPosSize(size.X, value, size.Width, size.Height, PosSize.Y)

    @visible.setter
    def visible(self, value: bool) -> None:
        with contextlib.suppress(Exception):
            model = cast(Any, self.get_view_ctl().getModel())
            if hasattr(model, "EnableVisible"):
                model.EnableVisible = value

    @property
    def name(self) -> str:
        """Gets the name for the control model"""
        try:
            model = cast(Any, self.get_view_ctl().getModel())
            return model.Name
        except Exception:
            return ""

    # endregion Properties


class CtlListenerBase(
    CtlBase,
    FocusEvents,
    KeyEvents,
    MouseEvents,
    MouseMotionEvents,
    PaintEvents,
    WindowEvents,
):
    """Dialog Control Listener Base Class"""

    # region Dunder Methods
    def __init__(self, ctl: Any) -> None:
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # The Events callback methods are invoked when any event is added or removed.
        FocusEvents.__init__(self, trigger_args=generic_args, cb=self._on_focus_listener_add_remove)
        KeyEvents.__init__(self, trigger_args=generic_args, cb=self._on_key_events_listener_add_remove)
        MouseEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_listener_add_remove)
        MouseMotionEvents.__init__(self, trigger_args=generic_args, cb=self._on_mouse_motion_listener_add_remove)
        PaintEvents.__init__(self, trigger_args=generic_args, cb=self._on_paint_listener_add_remove)
        WindowEvents.__init__(self, trigger_args=generic_args, cb=self._on_window_event_listener_add_remove)

    # endregion Dunder Methods

    # region Lazy Listeners

    # Listeners such as mouse, mouse motion, focus, key, paint, and window are added lazily.
    # Some listeners such as mouse motion may be expensive.
    # Each time the mouse moves, the listener is invoked.
    # By lazy loading listeners are only added when needed.
    # For example:
    #     ctl_button_ok.add_event_mouse_entered(on_mouse_entered)
    # This would call _on_mouse_listener_add_remove() below, which would add the mouse listener to the class in a lazy manor.

    def _on_mouse_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        view = cast(Any, self.get_view_ctl())
        view.addMouseListener(self.events_listener_mouse)
        self._add_listener(key)

    def _on_mouse_motion_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        view = cast(Any, self.get_view_ctl())
        view.addMouseMotionListener(self.events_listener_mouse_motion)
        self._add_listener(key)

    def _on_focus_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        view = cast(Any, self.get_view_ctl())
        view.addFocusListener(self.events_listener_focus)
        self._add_listener(key)

    def _on_key_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        view = cast(Any, self.get_view_ctl())
        view.addKeyListener(self.events_listener_key)
        self._add_listener(key)

    def _on_paint_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        view = cast(Any, self.get_view_ctl())
        view.addPaintListener(self.events_listener_paint)
        self._add_listener(key)

    def _on_window_event_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        view = cast(Any, self.get_view_ctl())
        view.addWindowListener(self.events_listener_window)
        self._add_listener(key)

    # endregion Lazy Listeners
