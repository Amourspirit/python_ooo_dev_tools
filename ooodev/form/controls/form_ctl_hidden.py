from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno

from ooodev.utils.kind.form_component_kind import FormComponentKind

from ooodev.form.controls.form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.component import HiddenControl as ControlModel  # service
    from com.sun.star.awt import UnoControl as ControlView  # service
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlHidden(FormCtlBase):
    """``com.sun.star.form.component.HiddenControl`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.HiddenControl`` service.
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
