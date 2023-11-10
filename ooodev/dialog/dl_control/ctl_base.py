from __future__ import annotations
import contextlib
from typing import Any
import uno
import unohelper

from com.sun.star.beans import XPropertySet

from ooodev.adapter.adapter_base import GenericArgs
from ooodev.adapter.awt.focus_events import FocusEvents
from ooodev.adapter.awt.key_events import KeyEvents
from ooodev.adapter.awt.mouse_events import MouseEvents
from ooodev.adapter.awt.mouse_motion_events import MouseMotionEvents
from ooodev.adapter.awt.paint_events import PaintEvents
from ooodev.adapter.awt.window_events import WindowEvents

# from ooodev.adapter.beans.properties_change_events import PropertiesChangeEvents
# from ooodev.adapter.beans.property_change_events import PropertyChangeEvents
# from ooodev.adapter.beans.vetoable_change_events import VetoableChangeEvents
from ooodev.utils import lo as mLo


class CtlBase(
    unohelper.Base,
    FocusEvents,
    KeyEvents,
    MouseEvents,
    MouseMotionEvents,
    PaintEvents,
    WindowEvents,
):
    """Dialog Control Base Class"""

    def __init__(self, ctl: Any) -> None:
        self._ctl_view = ctl
        generic_args = self._get_generic_args()
        FocusEvents.__init__(self, trigger_args=generic_args)
        KeyEvents.__init__(self, trigger_args=generic_args)
        MouseEvents.__init__(self, trigger_args=generic_args)
        MouseMotionEvents.__init__(self, trigger_args=generic_args)
        PaintEvents.__init__(self, trigger_args=generic_args)
        WindowEvents.__init__(self, trigger_args=generic_args)
        view = self.get_view_ctl()
        view.addMouseListener(self.events_listener_mouse)
        view.addMouseMotionListener(self.events_listener_mouse_motion)
        view.addFocusListener(self.events_listener_focus)
        view.addKeyListener(self.events_listener_key)
        view.addPaintListener(self.events_listener_paint)
        view.addWindowListener(self.events_listener_window)

    def _get_generic_args(self) -> GenericArgs:
        try:
            return self.__generic_args
        except AttributeError:
            self.__generic_args = GenericArgs(control_src=self)
            return self.__generic_args

    def __getattr__(self, name: str) -> Any:
        if hasattr(self._ctl_view, name):
            return getattr(self._ctl_view, name)
        raise AttributeError(name)

    def get_view_ctl(self) -> Any:
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

    @property
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        model = self.get_view_ctl().getModel()
        return model.Enabled if hasattr(model, "Enabled") else True

    @enabled.setter
    def enabled(self, value: bool) -> None:
        model = self.get_view_ctl().getModel()
        if hasattr(model, "Enabled"):
            model.Enabled = value

    @property
    def visible(self) -> bool:
        """Gets/Sets the visible state for the control"""
        try:
            model = self.get_view_ctl().getModel()
            return model.EnableVisible
        except Exception:
            return True

    @visible.setter
    def visible(self, value: bool) -> None:
        with contextlib.suppress(Exception):
            model = self.get_view_ctl().getModel()
            if hasattr(model, "EnableVisible"):
                model.EnableVisible = value

    @property
    def name(self) -> str:
        """Gets the name for the control model"""
        try:
            model = self.get_view_ctl().getModel()
            return model.Name
        except Exception:
            return ""
