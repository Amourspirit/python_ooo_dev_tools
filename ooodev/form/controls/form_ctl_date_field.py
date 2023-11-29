from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import datetime
from com.sun.star.awt import XControl

from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.date_format_kind import DateFormatKind
from ooodev.utils.kind.form_component_kind import FormComponentKind

from .form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.form.component import DateField as ControlModel  # service
    from com.sun.star.form.control import DateField as ControlView  # service
    from ooodev.events.args.listener_event_args import ListenerEventArgs


class FormCtlDateField(FormCtlBase, SpinEvents, TextEvents, ResetEvents):
    """``com.sun.star.form.component.DateField`` control"""

    def __init__(self, ctl: XControl) -> None:
        FormCtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        SpinEvents.__init__(self, trigger_args=generic_args, cb=self._on_spin_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        ResetEvents.__init__(self, trigger_args=generic_args, cb=self._on_reset_add_remove)

    # region Lazy Listeners
    def _on_spin_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addSpinListener(self.events_listener_spin)
        event.remove_callback = True

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
        event.remove_callback = True

    def _on_reset_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addResetListener(self.events_listener_reset)
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
        return FormComponentKind.DATE_FIELD

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
    def date(self) -> datetime.date:
        """Gets/Sets the date"""
        return DateUtil.uno_date_to_date(self.get_model().Date)

    @date.setter
    def date(self, value: datetime.date) -> None:
        self.get_model().Date = DateUtil.date_to_uno_date(value)

    @property
    def date_format(self) -> DateFormatKind:
        """Gets/Sets the format"""
        return DateFormatKind(self.get_model().DateFormat)

    @date_format.setter
    def date_format(self, value: DateFormatKind) -> None:
        self.get_model().DateFormat = value.value

    @property
    def date_max(self) -> datetime.date:
        """Gets/Sets the min date"""
        return DateUtil.uno_date_to_date(self.get_model().DateMax)

    @date_max.setter
    def date_max(self, value: datetime.date) -> None:
        self.get_model().DateMax = DateUtil.date_to_uno_date(value)

    @property
    def date_min(self) -> datetime.date:
        """Gets/Sets the min date"""
        return DateUtil.uno_date_to_date(self.get_model().DateMin)

    @date_min.setter
    def date_min(self, value: datetime.date) -> None:
        self.get_model().DateMin = DateUtil.date_to_uno_date(value)

    @property
    def dropdown(self) -> bool:
        """Gets/Sets the if the control has a dropdown."""
        return self.get_model().Dropdown

    @dropdown.setter
    def dropdown(self, value: bool) -> None:
        self.get_model().Dropdown = value

    @property
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        return self.get_model().Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.get_model().Enabled = value

    @property
    def help_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.get_model().HelpText = value

    @property
    def help_url(self) -> str:
        """Gets/Sets the help url"""
        return self.get_model().HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.get_model().HelpURL = value

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def printable(self) -> bool:
        """Gets/Sets the printable property"""
        return self.get_model().Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.get_model().Printable = value

    @property
    def read_only(self) -> bool:
        """Gets/Sets the read-only property"""
        return self.get_model().ReadOnly

    @read_only.setter
    def read_only(self, value: bool) -> None:
        """Sets the read-only property"""
        self.get_model().ReadOnly = value

    @property
    def spin(self) -> bool:
        """Gets/Sets if the control has a spin button"""
        return self.get_model().Spin

    @spin.setter
    def spin(self, value: bool) -> None:
        self.get_model().Spin = value

    @property
    def step(self) -> int:
        """Gets/Sets the step"""
        return self.get_model().Step

    @step.setter
    def step(self, value: int) -> None:
        self.get_model().Step = value

    @property
    def strict_format(self) -> bool:
        """Gets/Sets the strict format"""
        return self.get_model().StrictFormat

    @strict_format.setter
    def strict_format(self, value: bool) -> None:
        self.get_model().StrictFormat = value

    @property
    def tab_stop(self) -> bool:
        """Gets/Sets the tab stop property"""
        return self.model.Tabstop

    @tab_stop.setter
    def tab_stop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.get_model().Text = value

    @property
    def tip_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @tip_text.setter
    def tip_text(self, value: str) -> None:
        self.get_model().HelpText = value

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
