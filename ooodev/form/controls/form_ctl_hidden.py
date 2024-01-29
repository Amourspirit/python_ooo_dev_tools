from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno

from ooodev.utils.kind.form_component_kind import FormComponentKind

from .form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.component import HiddenControl as ControlModel  # service
    from com.sun.star.awt import UnoControl as ControlView  # service
    from ooodev.utils.inst.lo.lo_inst import LoInst


class FormCtlHidden(FormCtlBase):
    """``com.sun.star.form.component.HiddenControl`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        FormCtlBase.__init__(self, ctl=ctl, lo_inst=lo_inst)

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
        return FormComponentKind.HIDDEN_CONTROL

    # endregion Overrides

    # region Properties
    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
