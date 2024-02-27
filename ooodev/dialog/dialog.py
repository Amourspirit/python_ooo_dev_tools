from __future__ import annotations
from typing import cast, TYPE_CHECKING

import uno
from com.sun.star.awt import XControl
from com.sun.star.awt import XControlModel
from com.sun.star.beans import XPropertySet

# from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.dialog.partial.dialog_controls_partial import DialogControlsPartial
from ooodev.dialog.partial.dialogs_partial import DialogsPartial
from ooodev.dialog.dl_control.ctl_dialog import CtlDialog
from ooodev.loader import lo as mLo
from ooodev.exceptions import ex as mEx
from ooodev.utils.context.lo_context import LoContext

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog  # service
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.units.unit_obj import UnitT


class Dialog(CtlDialog, DialogControlsPartial, DialogsPartial):
    """Class for creating a Dialog. Multiple Document are supported"""

    def __init__(
        self,
        title: str,
        x: int | UnitT = -1,
        y: int | UnitT = -1,
        width: int | UnitT = -1,
        height: int | UnitT = -1,
        *,
        lo_inst: LoInst | None = None,
        dialog: UnoControlDialog | None = None,
    ) -> None:
        """
        Dialog Constructor

        Args:
            title (str): Dialog title.
            x (int, UnitT): The x-coordinate of the window. In ``1/100 mm`` or ``UnitT``. If ``-1``, the dialog Position is not set.
                Default is ``-1``.
            y (int, UnitT): The y-coordinate of the window. In ``1/100 mm`` or ``UnitT``. If ``-1``, the dialog Position is not set.
                Default is ``-1``.
            width (int, UnitT): The width of the window. In ``1/100 mm`` or ``UnitT``. If ``-1``, the dialog Size is not set.
                Default is ``-1``.
            height (int, UnitT): The height of the window. In ``1/100 mm`` or ``UnitT``. If ``-1``, the dialog Size is not set.
                Default is ``-1``.
            lo_inst (LoInst, optional): Lo Instance. Used when creating multiple documents. Defaults to ``None``.
            ctl (UnoControlDialog): Dialog Control. If ``None`` then an new Dialog is created. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        if dialog is None:
            dialog = self.__create_dialog(lo_inst=lo_inst, x=x, y=y, width=width, height=height, title=title)
        with LoContext(inst=lo_inst):
            CtlDialog.__init__(self, ctl=dialog)
        # LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DialogControlsPartial.__init__(self, dialog_ctl=dialog, lo_inst=self.lo_inst)
        DialogsPartial.__init__(self, dialog_ctl=dialog, lo_inst=self.lo_inst)

    def __create_dialog(
        self,
        lo_inst: LoInst,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        title: str,
    ) -> UnoControlDialog:
        """
        Creates a dialog.

        Args:
            title (str): Dialog title.
            x (int, UnitT): The x-coordinate of the window. In ``1/100 mm`` or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): The y-coordinate of the window. In ``1/100 mm`` or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): The width of the window. In ``1/100 mm`` or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT): The height of the window. In ``1/100 mm`` or ``UnitT``. If ``-1``, the dialog Size is not set.

        Raises:
            DialogError: If unable to create dialog.

        Returns:
            CtlDialog: Control.

        See Also:
            `API UnoControlDialogModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogModel.html>`_
        """
        # sourcery skip: raise-specific-error
        # pylint: disable=protected-access
        try:
            try:
                x_arg = cast(int, x.get_value_mm100())  # type: ignore
            except AttributeError:
                x_arg = cast(int, x)
            try:
                y_arg = cast(int, y.get_value_mm100())  # type: ignore
            except AttributeError:
                y_arg = cast(int, y)
            try:
                width_arg = cast(int, width.get_value_mm100())  # type: ignore
            except AttributeError:
                width_arg = cast(int, width)
            try:
                height_arg = cast(int, height.get_value_mm100())  # type: ignore
            except AttributeError:
                height_arg = cast(int, height)
            dialog_ctrl = cast(
                "UnoControlDialog",
                lo_inst.create_instance_mcf(XControl, "com.sun.star.awt.UnoControlDialog", raise_err=True),
            )
            control_model = lo_inst.create_instance_mcf(
                XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True
            )
            dialog_ctrl.setModel(control_model)
            model = dialog_ctrl.getModel()

            ctl_props = lo_inst.qi(XPropertySet, model, True)
            ctl_props.setPropertyValue("Title", title)
            ctl_props.setPropertyValue("Name", "OfficeDialog")

            ctl_props.setPropertyValue("Step", 0)
            ctl_props.setPropertyValue("Moveable", True)
            ctl_props.setPropertyValue("TabIndex", 0)
            self._DialogControlsPartial_dialogs_class._set_size_pos(dialog_ctrl, x_arg, y_arg, width_arg, height_arg)
            return dialog_ctrl
        except Exception as e:
            raise mEx.DialogError(f"Could not create dialog control: {e}") from e
