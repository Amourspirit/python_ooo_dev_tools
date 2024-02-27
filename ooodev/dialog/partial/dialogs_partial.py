# coding: utf-8
# pylint: disable=too-many-lines
# region Imports
from __future__ import annotations
from typing import TYPE_CHECKING, Any, List, Tuple, Type

# pylint: disable=unused-import
import uno
from com.sun.star.awt import XControl
from com.sun.star.awt import XControlContainer
from com.sun.star.awt import XTopWindow
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNameContainer

from ooodev.mock import mock_g
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.dialog.dl_control.ctl_base import DialogControlBase


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog  # service
    from ooodev.loader.inst.lo_inst import LoInst

    # Avoid circular import by creating a property in class instance for Dialogs
    from ooodev.dialog.dialogs import Dialogs
    from ooodev.dialog.dialogs import ControlT
    from ooodev.dialog.dl_control import CtlRadioButton
# endregion Imports


class DialogsPartial:
    def __init__(self, dialog_ctl: UnoControlDialog, lo_inst: LoInst | None = None) -> None:
        """
        DialogsPartial Constructor.

        Args:
            dialog_ctl (UnoControlDialog): Main Dialog Window Control.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__ctl = dialog_ctl

    def find_controls(self, control_type: Type[ControlT], dialog_ctrl: XControl | None = None) -> List[ControlT]:
        """
        Finds controls by type.

        Args:
            control_type (ControlT): Control type.
            dialog_ctrl (XControl, optional): Control. Defaults to Instance control.

        Returns:
            List[ControlT]: List of controls.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogsPartial_dialogs.find_controls(dialog_ctrl=dialog_ctrl, control_type=control_type)
        return result

    def create_peer(self) -> None:
        """
        Creates a dialog peer.
        """
        with LoContext(inst=self.__lo_inst):
            _ = self._DialogsPartial_dialogs.create_dialog_peer(self.__ctl)

    def find_radio_siblings(self, radio_button: str, dialog_ctrl: XControl | None = None) -> List[CtlRadioButton]:
        """
        Given the name of the first radio button of a group, return all the controls of the group.

        For dialogs, radio buttons are considered of the same group when their tab indexes are contiguous.

        Args:
            radio_button (str): Specifies the exact name of the 1st radio button of the group.
            dialog_ctrl (XControl, optional): Control. Defaults to Instance control.

        Returns:
            List[CtlRadioButton]: List of the names of the 1st and the next radio buttons.
            belonging to the same group in their tab index order. does not include the first button.

        See Also:
            :py:meth:`~.dialogs.Dialogs.get_radio_group_value`.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            results = self._DialogsPartial_dialogs.find_radio_siblings(
                dialog_ctrl=dialog_ctrl, radio_button=radio_button
            )
        return results

    def get_control_class_id(self, dialog_ctrl: XControl) -> str:
        """
        Gets control class id.

        Args:
            dialog_ctrl (XControl): Control.

        Returns:
            str: class id.
        """
        return self._DialogsPartial_dialogs.get_control_class_id(dialog_ctrl)

    def get_control_name(self, dialog_ctrl: XControl) -> str:
        """
        Get the name of a control.

        Args:
            dialog_ctrl (XControl): Control.

        Returns:
            str: control name.
        """
        return self._DialogsPartial_dialogs.get_control_name(dialog_ctrl)

    def get_control_props(self, control_model: Any) -> XPropertySet:
        """
        Gets property set for a control model.

        |lo_safe|

        Args:
            control_model (Any): control model.

        Returns:
            XPropertySet: Property set.
        """
        return self.__lo_inst.qi(XPropertySet, control_model, True)

    def control_factory(self, dialog_ctrl: XControl) -> DialogControlBase | None:
        """
        Gets a control as a ``DialogControlBase`` control.

        Args:
            dialog_ctrl (XControl): Control.

        Returns:
            DialogControlBase | None: Returns a ``DialogControlBase`` such as ``CtlButton`` or ``CtlCheckBox`` if found, else ``None``
        """
        with LoContext(inst=self.__lo_inst):
            result = self._DialogsPartial_dialogs.get_dialog_control_instance(dialog_ctrl)
        return result

    def get_controls_arr(self, dialog_ctrl: XControl) -> Tuple[XControl, ...]:
        """
        Gets all controls for a given control.

        Args:
            dialog_ctrl (XControl): control.

        Returns:
            Tuple[XControl, ...]: controls.
        """
        ctrl_con = self.__lo_inst.qi(XControlContainer, dialog_ctrl, True)
        return ctrl_con.getControls()

    def get_name_container(self, dialog_ctrl: XControl) -> XNameContainer:
        """
        Gets Name container from control.

        Args:
            dialog_ctrl (XControl): Dialog control.

        Returns:
            XNameContainer: Name Container.
        """
        return self.__lo_inst.qi(XNameContainer, dialog_ctrl.getModel(), True)

    def get_window(self, dialog_ctrl: XControl) -> XTopWindow:
        """
        Gets dialog window.

        Args:
            dialog_ctrl (XControl): dialog control.

        Returns:
            XTopWindow: Top window instance.
        """
        return self.__lo_inst.qi(XTopWindow, dialog_ctrl, True)

    def get_radio_group_value(self, dialog_ctrl: XControl, radio_button: str) -> List[CtlRadioButton]:
        """
        Get a radio button group. Similar to :py:meth:`~.dialogs.Dialogs.find_radio_siblings` but also includes first radio button.

        Args:
            dialog_ctrl (XControl): Control.
            radio_button (str): Name of the first radio button of the group.

        Returns:
            Any: Value of the selected radio button.

        See Also:
            :py:meth:`~.dialogs.Dialogs.find_radio_siblings`.
        """
        with LoContext(inst=self.__lo_inst):
            result = self._DialogsPartial_dialogs.get_radio_group_value(dialog_ctrl, radio_button)
        return result

    @property
    def _DialogsPartial_dialogs(self) -> Type[Dialogs]:
        try:
            # avoid circular import.
            return self._DialogsPartial_dialogs_class_instance
        except AttributeError:
            # pylint: disable=import-outside-toplevel
            from ooodev.dialog.dialogs import Dialogs

            self._DialogsPartial_dialogs_class_instance = Dialogs
        return self._DialogsPartial_dialogs_class_instance


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dialogs import Dialogs
