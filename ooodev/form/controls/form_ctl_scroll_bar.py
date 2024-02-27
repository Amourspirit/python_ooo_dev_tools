from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib

from ooodev.adapter.awt.adjustment_events import AdjustmentEvents
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.utils.kind.orientation_kind import OrientationKind as OrientationKind

from ooodev.form.controls.form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.component import ScrollBar as ControlModel  # service
    from com.sun.star.awt import UnoControlScrollBar as ControlView  # service
    from ooodev.events.args.listener_event_args import ListenerEventArgs
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlScrollBar(FormCtlBase, AdjustmentEvents, ResetEvents):
    """``com.sun.star.form.component.ScrollBar`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.ScrollBar`` service.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to ``None``.

        Returns:
            None:

        Note:
            If the :ref:`LoContext <ooodev.utils.context.lo_context.LoContext>` manager is use before this class is instantiated,
            then the Lo instance will be set using the current Lo instance. That the context manager has set.
            Generally speaking this means that there is no need to set ``lo_inst`` when instantiating this class.

        See Also:
            :ref:`ooodev.form.Forms`.
        """
        FormCtlBase.__init__(self, ctl=ctl, lo_inst=lo_inst)
        generic_args = self._get_generic_args()
        ResetEvents.__init__(self, trigger_args=generic_args, cb=self._on_reset_add_remove)

    # region Lazy Listeners

    def _on_reset_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addResetListener(self.events_listener_reset)
        event.remove_callback = True

    def _on_adjustment_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        # UnoControlScrollBar as view is not documented. So Not sure about this one
        with contextlib.suppress(AttributeError):
            self.view.addAdjustmentListener(self.events_listener_adjustment)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    if TYPE_CHECKING:
        # override the methods to provide type hinting
        def get_view(self) -> ControlView:
            """Gets the view of this control"""
            return cast("ControlView", super().get_view())

        def get_model(self) -> ControlModel:
            """Gets the model for this control"""
            return cast("ControlModel", super().get_model())

    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        return FormComponentKind.SCROLL_BAR

    # endregion Overrides

    # region Properties
    @property
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.Border = value.value

    @property
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

    @property
    def help_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def help_url(self) -> str:
        """Gets/Sets the help url"""
        return self.model.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.model.HelpURL = value

    @property
    def min_value(self) -> int:
        """Gets the minimum value of the scroll bar"""
        return self.model.ScrollValueMin

    @min_value.setter
    def min_value(self, value: int) -> None:
        self.model.ScrollValueMin = value

    @property
    def max_value(self) -> int:
        """Gets the maximum value of the scroll bar"""
        return self.model.ScrollValueMax

    @max_value.setter
    def max_value(self, value: int) -> None:
        self.model.ScrollValueMax = value

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def orientation(self) -> OrientationKind:
        """Gets or sets the orientation of the scroll bar"""
        return OrientationKind(self.model.Orientation)

    @orientation.setter
    def orientation(self, value: OrientationKind) -> None:
        self.model.Orientation = value.value

    @property
    def printable(self) -> bool:
        """Gets/Sets the printable property"""
        return self.model.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.model.Printable = value

    @property
    def step(self) -> int:
        """Gets/Sets the step"""
        return self.model.Step

    @step.setter
    def step(self, value: int) -> None:
        self.model.Step = value

    @property
    def tab_stop(self) -> bool:
        """Gets/Sets the tab stop property"""
        return self.model.Tabstop

    @tab_stop.setter
    def tab_stop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def tip_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @tip_text.setter
    def tip_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def value(self) -> int:
        """
        Gets or sets the current value of the scroll bar.

        Return ``-1`` if the scroll bar does not support this.
        """
        # not sure if this scroll bar supports this.
        with contextlib.suppress(AttributeError):
            return self.view.getValue()
        return -1

    @value.setter
    def value(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.view.setValue(value)

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
