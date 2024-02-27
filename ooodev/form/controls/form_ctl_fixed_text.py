from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl

from ooodev.adapter.awt.text_events import TextEvents
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.utils import info as mInfo

from ooodev.form.controls.form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.form import DataAwareControlModel
    from com.sun.star.form.component import FixedText as ControlModel  # service
    from com.sun.star.form.control import TextField as ControlView  # service
    from ooodev.events.args.listener_event_args import ListenerEventArgs
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlFixedText(FormCtlBase, TextEvents):
    """``com.sun.star.form.component.FixedText`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.FixedText`` service.
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
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # region Lazy Listeners

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
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
        return FormComponentKind.FIXED_TEXT

    def _get_tab_index(self) -> int:
        """
        Gets the tab index.

        Fixed Text controls do not have a tab index, so this always returns -1
        """
        return -1

    def _set_tab_index(self, value: int) -> None:
        """
        Sets the tab index.

        Fixed Text controls do not have a tab index, so this does nothing
        """
        pass

    # endregion Overrides

    # region Methods
    def bind_to_control(self, ctl: FormCtlBase) -> None:
        """Binds this control to the given control"""
        model = ctl.get_model()
        if not mInfo.Info.support_service(model, "com.sun.star.form.DataAwareControlModel"):
            return
        data_model = cast("DataAwareControlModel", model)
        data_model.DataField = self.name
        data_model.LabelControl = self.get_property_set()
        return

    # endregion Methods

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
    def label(self) -> str:
        """Gets/Sets the label"""
        return self.model.Label

    @label.setter
    def label(self, value: str) -> None:
        self.model.Label = value

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def multi_line(self) -> bool:
        """Gets/Sets the multi line"""
        return self.model.MultiLine

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        self.model.MultiLine = value

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
    def tip_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @tip_text.setter
    def tip_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
