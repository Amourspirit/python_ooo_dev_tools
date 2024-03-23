# coding: utf-8
# pylint: disable=too-many-lines
# region Imports
from __future__ import annotations
import datetime
from typing import TYPE_CHECKING, Any, Iterable, Type

# pylint: disable=unused-import
import uno

from com.sun.star.awt import XControl

from ooo.dyn.awt.image_scale_mode import ImageScaleModeEnum
from ooo.dyn.awt.pos_size import PosSize as PosSize
from ooo.dyn.awt.push_button_type import PushButtonType
from ooo.dyn.style.vertical_alignment import VerticalAlignment
from ooo.dyn.awt.pos_size import PosSizeEnum as PosSizeEnum

from ooodev.mock import mock_g
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.kind.align_kind import AlignKind
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.kind.date_format_kind import DateFormatKind
from ooodev.utils.kind.horz_ver_kind import HorzVertKind
from ooodev.utils.kind.orientation_kind import OrientationKind
from ooodev.utils.kind.time_format_kind import TimeFormatKind
from ooodev.utils.kind.tri_state_kind import TriStateKind
from ooodev.dialog.dl_control.ctl_button import CtlButton
from ooodev.dialog.dl_control.ctl_check_box import CtlCheckBox
from ooodev.dialog.dl_control.ctl_combo_box import CtlComboBox
from ooodev.dialog.dl_control.ctl_currency_field import CtlCurrencyField
from ooodev.dialog.dl_control.ctl_date_field import CtlDateField
from ooodev.dialog.dl_control.ctl_file import CtlFile
from ooodev.dialog.dl_control.ctl_fixed_line import CtlFixedLine
from ooodev.dialog.dl_control.ctl_fixed_text import CtlFixedText
from ooodev.dialog.dl_control.ctl_formatted_field import CtlFormattedField
from ooodev.dialog.dl_control.ctl_grid import CtlGrid
from ooodev.dialog.dl_control.ctl_group_box import CtlGroupBox
from ooodev.dialog.dl_control.ctl_hyperlink_fixed import CtlHyperlinkFixed
from ooodev.dialog.dl_control.ctl_image import CtlImage
from ooodev.dialog.dl_control.ctl_list_box import CtlListBox
from ooodev.dialog.dl_control.ctl_numeric_field import CtlNumericField
from ooodev.dialog.dl_control.ctl_pattern_field import CtlPatternField
from ooodev.dialog.dl_control.ctl_progress_bar import CtlProgressBar
from ooodev.dialog.dl_control.ctl_radio_button import CtlRadioButton
from ooodev.dialog.dl_control.ctl_scroll_bar import CtlScrollBar
from ooodev.dialog.dl_control.ctl_spin_button import CtlSpinButton
from ooodev.dialog.dl_control.ctl_tab_page import CtlTabPage
from ooodev.dialog.dl_control.ctl_tab_page_container import CtlTabPageContainer
from ooodev.dialog.dl_control.ctl_text_edit import CtlTextEdit
from ooodev.dialog.dl_control.ctl_time_field import CtlTimeField
from ooodev.dialog.dl_control.ctl_tree import CtlTree


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDialog  # service
    from ooodev.utils.type_var import PathOrStr
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.units.unit_obj import UnitT

    # Avoid circular import by creating a property in class instance for Dialogs
    from ooodev.dialog.dialogs import Dialogs
# endregion Imports


class DialogControlsPartial:
    """Partial Class for Dialog Controls."""

    def __init__(self, dialog_ctl: UnoControlDialog, lo_inst: LoInst | None = None) -> None:
        """
        Dialog Controls Constructor.

        Args:
            dialog_ctl (UnoControlDialog): Main Dialog Window Control.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__ctl = dialog_ctl

    # region Insert Controls
    # region Insert Button
    def insert_button(
        self,
        *,
        label: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        btn_type: PushButtonType | None = None,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlButton:
        """
        Insert Button Control.

        Args:
            label (str): Button Label.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            btn_type (PushButtonType | None, optional): Type of Button.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create button control.

        Returns:
            CtlButton: Button control.

        Hint:
            - ``PushButtonType`` can be imported from ``ooo.dyn.awt.push_button_type``.

        See Also:
            `API UnoControlButtonModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlButtonModel.html>`_
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_button(
                dialog_ctrl,
                label=label,
                x=x,
                y=y,
                width=width,
                height=height,
                btn_type=btn_type,
                name=name,
                **props,
            )
        return result

    # endregion Insert Button

    # region Insert CheckBox
    def insert_check_box(
        self,
        *,
        label: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 8,
        tri_state: bool = True,
        state: TriStateKind = TriStateKind.CHECKED,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlCheckBox:
        """
        Inserts a check box control.

        Args:
            label (str): Checkbox label text.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int,, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``8``. If ``-1``, the dialog Size is not set.
            tri_state (TriStateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``TriStateKind.CHECKED``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of Checkbox. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create checkbox control.

        Returns:
            CtlCheckBox: Check box control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``TriStateKind`` can be imported from ``ooodev.utils.kind.tri_state_kind``.

        See Also:
            `API UnoControlCheckBoxModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlCheckBoxModel.html>`_
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_check_box(
                dialog_ctrl,
                label=label,
                x=x,
                y=y,
                width=width,
                height=height,
                tri_state=tri_state,
                state=state,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert CheckBox

    # region Insert ComboBox

    def insert_combo_box(
        self,
        *,
        entries: Iterable[str],
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        max_text_len: int = 0,
        drop_down: bool = True,
        read_only: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlComboBox:
        """
        Insert a combo box control.

        Args:
            entries (Iterable[str]): Combo box entries.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            max_text_len (int, optional): Specifies the maximum character count, There's no limitation, if set to 0. Defaults to ``0``.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``True``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create combo box control.

        Returns:
            CtlComboBox: Combo box control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.

        See Also:
            `API UnoControlComboBoxModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlComboBoxModel.html>`_
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_combo_box(
                dialog_ctrl,
                entries=entries,
                x=x,
                y=y,
                width=width,
                height=height,
                max_text_len=max_text_len,
                drop_down=drop_down,
                read_only=read_only,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert ComboBox

    # region Insert Currency Field

    def insert_currency_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        value: float = 0.0,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlCurrencyField:
        """
        Inserts a currency control.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            value (float, optional): Control Value. Defaults to ``0.0``.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.

        Raises:
            DialogError: If unable to create currency field box control.

        Returns:
            CtlCurrencyField: Currency field control.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_currency_field(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                value=value,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                increment=increment,
                accuracy=accuracy,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Currency Field

    # region Insert Date Field

    def insert_date_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        date_value: datetime.datetime | None = None,
        min_date: datetime.datetime = datetime.datetime(1900, 1, 1, 0, 0, 0, 0),
        max_date: datetime.datetime = datetime.datetime(2200, 12, 31, 0, 0, 0, 0),
        drop_down: bool = True,
        date_format: DateFormatKind = DateFormatKind.SYSTEM_SHORT,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlDateField:
        """
        Create a new control of type DateField in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            date_value (datetime.datetime | None, optional): Specifics control datetime. Defaults to ``None``.
            min_date (datetime.datetime, optional): Specifics control min datetime. Defaults to ``datetime(1900, 1, 1, 0, 0, 0, 0)``.
            max_date (datetime.datetime, optional): Specifics control Min datetime. Defaults to ``datetime(2200, 12, 31, 0, 0, 0, 0)``.
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to ``True``.
            date_format (DateFormatKind, optional): Date format. Defaults to ``DateFormatKind.SYSTEM_SHORT``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create date field control.

        Returns:
            CtlDateField: Date field control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``DateFormatKind`` can be imported from ``ooodev.utils.kind.date_format_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_date_field(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                date_value=date_value,
                min_date=min_date,
                max_date=max_date,
                drop_down=drop_down,
                date_format=date_format,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Date Field

    # region Insert File Control

    def insert_file_control(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlFile:
        """
        Create a new control of type FileControl in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create file control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.

        Returns:
            CtlFile: File Control.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_file_control(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert File Control

    # region Insert Fixed Line

    def insert_fixed_line(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 1,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlFixedLine:
        """
        Create a new control of type Fixed Line in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``1``. If ``-1``, the dialog Size is not set.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create fixed line control.

        Returns:
            CtlFixedLine: Fixed Line Control.

        Hint:
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_fixed_line(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                orientation=orientation,
                name=name,
                **props,
            )
        return result

    # endregion Insert Fixed Line

    # region Insert Formatted Field

    def insert_formatted_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        value: float | None = None,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlFormattedField:
        """
        Create a new control of type FormattedField in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            value (float, optional): Control Value. Defaults to ``0.0``.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create formatted field control.

        Returns:
            CtlFormattedField: Formatted Field.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_formatted_field(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                value=value,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Formatted Field

    # region Insert Group Box

    def insert_group_box(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        label: str = "",
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlGroupBox:
        """
        Create a new control of type GroupBox in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            label (str, optional): Group box label.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create group box control.

        Returns:
            CtlGroupBox: Group box Control.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_group_box(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                label=label,
                name=name,
                **props,
            )
        return result

    # endregion Insert Group Box

    # region Insert Hyperlink

    def insert_hyperlink(
        self,
        *,
        label: str,
        url: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        align: AlignKind = AlignKind.LEFT,
        vert_align: VerticalAlignment = VerticalAlignment.TOP,
        multiline: bool = False,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlHyperlinkFixed:
        """
        Create a new control of type Hyperlink in the actual dialog.

        Args:
            label (str): Hyperlink label.
            url (str): Hyperlink URL.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            align (AlignKind, optional): Horizontal alignment. Defaults to ``AlignKind.LEFT``.
            vert_align (VerticalAlignment, optional): Vertical alignment. Defaults to ``VerticalAlignment.TOP``.
            multiline (bool, optional): Specifies if the control can display multiple lines of text. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create Hyperlink control.

        Returns:
            CtlHyperlinkFixed: Hyperlink Control.

        Hint:
            - ``AlignKind`` can be imported from ``ooodev.utils.kind.align_kind``.
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_hyperlink(
                dialog_ctrl,
                label=label,
                url=url,
                x=x,
                y=y,
                width=width,
                height=height,
                align=align,
                vert_align=vert_align,
                multiline=multiline,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Hyperlink

    # region Insert Image Control

    def insert_image_control(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        border: BorderKind = BorderKind.BORDER_3D,
        scale: int | ImageScaleModeEnum = ImageScaleModeEnum.NONE,
        image_url: PathOrStr = "",
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlImage:
        """
        Create a new control of type ImageControl in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, , UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            image_url (PathOrStr, optional): Image URL. When setting the value it can be a string or a Path object.
                If a string is passed it can be a URL or a path to a file.
                Value such as ``file:///path/to/image.png`` and ``/path/to/image.png`` are valid.
                Relative paths are supported.
            scale (int | ImageScaleModeEnum, optional): Image scale mode. Defaults to ``ImageScaleModeEnum.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create image control.

        Returns:
            CtlImage: Image Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``ImageScaleModeEnum`` can be imported from ``ooo.dyn.awt.image_scale_mode``.
            - ``VerticalAlignment`` can be imported from ``ooo.dyn.style.vertical_alignment``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_image_control(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                border=border,
                scale=scale,
                image_url=image_url,
                name=name,
                **props,
            )
        return result

    # endregion Insert Image Control

    # region Insert Label

    def insert_label(
        self,
        *,
        label: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlFixedText:
        """
        Insert a Fixed Text into a control.

        Args:
            label (str): Contents of label.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Default ``20``. If ``-1``, the dialog Size is not set.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create Fixed Text.

        Returns:
            CtlFixedText: Fixed Text Control.

        See Also:
            `API UnoControlFixedTextModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlFixedTextModel.html>`_
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_label(
                dialog_ctrl,
                label=label,
                x=x,
                y=y,
                width=width,
                height=height,
                name=name,
                **props,
            )
        return result

    # endregion Insert Label

    # region Insert List Box
    def insert_list_box(
        self,
        *,
        entries: Iterable[str],
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 100,
        drop_down: bool = False,
        read_only: bool = False,
        line_count: int = 5,
        multi_select: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlListBox:
        """
        Insert a list box control.

        Args:
            entries (Iterable[str]): List box entries.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``100``. If ``-1``, the dialog Size is not set.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``False``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            line_count (int, optional): Specifies the number of lines to display. Defaults to ``5``.
            multi_select (int, optional): Specifies if multiple entries can be selected. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create list box control.

        Returns:
            CtlListBox: List box control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_list_box(
                dialog_ctrl,
                entries=entries,
                x=x,
                y=y,
                width=width,
                height=height,
                drop_down=drop_down,
                read_only=read_only,
                line_count=line_count,
                multi_select=multi_select,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert List Box

    # region Insert Numeric Field
    def insert_numeric_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        value: float | None = None,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlNumericField:
        """
        Create a new control of type NumericField in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Default ``20``. If ``-1``, the dialog Size is not set.
            value (float, optional): Control Value. Defaults to ``0.0``.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create numeric field control.

        Returns:
            CtlNumericField: Numeric Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_numeric_field(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                value=value,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                increment=increment,
                accuracy=accuracy,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Numeric Field

    # region Insert Password Field
    def insert_password_field(
        self,
        *,
        text: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlTextEdit:
        """
        Inserts a password field.

        Args:
            text (str): Text value.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create text field.

        Returns:
            CtlTextEdit: Text Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.

        See Also:
            :py:meth:`~.dialogs.self._DialogControlsPartial_dialogs_class.insert_text_field`.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_password_field(
                dialog_ctrl,
                text=text,
                x=x,
                y=y,
                width=width,
                height=height,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Password Field

    # region Insert Pattern Field
    def insert_pattern_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        edit_mask: str = "",
        literal_mask: str = "",
        border: BorderKind = BorderKind.BORDER_3D,
        dialog_ctrl: XControl | None = None,
        name: str = "",
        **props: Any,
    ) -> CtlPatternField:
        """
        Create a new control of type PatternField in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Default``20``. If ``-1``, the dialog Size is not set.
            edit_mask (str, optional): Specifies a character code that determines what the user may enter. Defaults to ``""``.
            literal_mask (str, optional): Specifies the initial values that are displayed in the pattern field. Defaults to ``""``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create pattern field control.

        Returns:
            CtlPatternField: Pattern Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_pattern_field(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                edit_mask=edit_mask,
                literal_mask=literal_mask,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Pattern Field

    # region Insert Progress Bar
    def insert_progress_bar(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        min_value: int = 0,
        max_value: int = 100,
        value: int = 0,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlProgressBar:
        """
        Create a new control of type Progress Control in the actual dialog.

        Args:
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT): Height in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            value (int, optional): The value initial value of the progress bar. Defaults to ``0``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create numeric field control.

        Returns:
            CtlProgressBar: Progress Bar Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_progress_bar(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                value=value,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Progress Bar

    # region Insert Radio Button
    def insert_radio_button(
        self,
        *,
        label: str,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        multiline: bool = False,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlRadioButton:
        """
        Create a new control of type RadioButton in the actual dialog.

        Args:
            label (str): Contents of label.
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Default ``20``. If ``-1``, the dialog Size is not set.
            multiline (bool, optional): Specifies if the control can display multiple lines of text. Defaults to ``False``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create radio button.

        Returns:
            CtlRadioButton: Radio button control.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_radio_button(
                dialog_ctrl,
                label=label,
                x=x,
                y=y,
                width=width,
                height=height,
                multiline=multiline,
                name=name,
                **props,
            )
        return result

    # endregion Insert Radio Button

    # region Insert Scroll Bar
    def insert_scroll_bar(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        min_value: int = 0,
        max_value: int = 100,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlScrollBar:
        """
        Create a new control of type ScrollBar in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT): Height in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create scroll bar control.

        Returns:
            CtlScrollBar: Scroll Bar Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_scroll_bar(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                orientation=orientation,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Scroll Bar

    # region Insert Spin Button
    def insert_spin_button(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        min_value: int = 0,
        max_value: int = 100,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlSpinButton:
        """
        Create a new control of type SpinButton in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT): Height in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create scroll bar control.

        Returns:
            CtlSpinButton: Scroll Bar Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_spin_button(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                orientation=orientation,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Spin Button

    # region Insert Tab Control

    def insert_tab_control(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 1,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        dialog_ctrl: XControl | None = None,
    ) -> CtlTabPageContainer:
        """
        Create a new control of type tab in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``1``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.

        Raises:
            DialogError: If unable to create tab control.

        Returns:
            CtlTabPageContainer: Tab Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.

        See Also:
            :py:meth:`~.dialogs.self._DialogControlsPartial_dialogs_class.insert_tab_page`.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_tab_control(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                border=border,
                name=name,
            )
        return result

    # endregion Insert Tab Control

    # region Insert Tab Page

    def insert_tab_page(
        self,
        *,
        tab_ctrl: CtlTabPageContainer,
        title: str,
        tab_position: int,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlTabPage:
        """
        Create a new control of type Tab Page in the actual tab control.

        Args:
            tab_ctrl (CtlTabPageContainer): Tab Container.
            title (str): Tab title.
            tab_position (int): Tab position.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create tab page control.

        Returns:
            CtlTabPage: Tab Page Control.

        See Also:
            :py:meth:`~.dialogs.self._DialogControlsPartial_dialogs_class.insert_tab_control`.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_tab_page(
                dialog_ctrl,
                tab_ctrl=tab_ctrl,
                title=title,
                tab_position=tab_position,
                name=name,
                **props,
            )
        return result

    # endregion Insert Tab Page

    # region Insert Table Control
    def insert_table_control(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        row_header: bool = True,
        col_header: bool = True,
        grid_lines: bool = False,
        scroll_bars: HorzVertKind = HorzVertKind.NONE,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlGrid:
        """
        Create a new control of type TableControl in the actual dialog.

        To add data to the table use :py:meth:`~.dialogs.self._DialogControlsPartial_dialogs_class.set_table_data`.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT): Height in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            row_header (bool, optional): Specifies if the control has a row header. Defaults to ``True``.
            col_header (bool, optional): Specifies if the control has a column header. Defaults to ``True``.
            grid_lines (bool, optional): Specifies if the control has grid lines. when True horizontal and vertical lines are painted between the grid cells. Defaults to ``False``.
            scroll_bars (HorzVertKind, optional): Specifies if the control has scroll bars. Scrollbars always appear dynamically when they are needed. Defaults to ``HorzVertKind.NONE``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create table control.

        Returns:
            CtlGrid: Table Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``HorzVertKind`` can be imported from ``ooodev.utils.kind.horz_ver_kind``.

        See Also:
            :py:meth:`~.dialogs.self._DialogControlsPartial_dialogs_class.set_table_data`.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_table_control(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                row_header=row_header,
                col_header=col_header,
                grid_lines=grid_lines,
                scroll_bars=scroll_bars,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Table Control

    # region Insert Text Field
    def insert_text_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        text: str = "",
        echo_char: str = "",
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlTextEdit:
        """
        Inserts a text Field.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            text (str, optional): Text value.
            echo_char (str, optional): Character used for masking. Must be a single character.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create text field.

        Returns:
            CtlTextEdit: Text Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.

        See Also:
            `API UnoControlEditModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlEditModel.html>`_
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_text_field(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                text=text,
                echo_char=echo_char,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Text Field

    # region Insert Time Field
    def insert_time_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 20,
        time_value: datetime.time | None = None,
        min_time: datetime.time = datetime.time(0, 0, 0, 0),
        max_time: datetime.time = datetime.time(23, 59, 59, 999_999),
        time_format: TimeFormatKind = TimeFormatKind.SHORT_24H,
        spin_button: bool = True,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlTimeField:
        """
        Create a new control of type TimeField in the actual dialog.

        Args:
            x (int, UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int, UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int, UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int, UnitT, optional): Height in Pixels or ``UnitT``. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            time_value (datetime.time | None, optional): Specifics the control time. Defaults to ``None``.
            min_time (datetime.time, optional): Specifics control min time. Defaults to ``time(0, 0, 0, 0)``.
            max_time (datetime.time, optional): Specifics control min time. Defaults to a ``time(23, 59, 59, 999_999)``.
            time_format (TimeFormatKind, optional): Date format. Defaults to ``TimeFormatKind.SHORT_24H``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``True``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create time field control.

        Returns:
            CtlTimeField: Time field control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``TimeFormatKind`` can be imported from ``ooodev.utils.kind.time_format_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_time_field(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                time_value=time_value,
                min_time=min_time,
                max_time=max_time,
                time_format=time_format,
                spin_button=spin_button,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Time Field

    # region Insert Tree Control
    def insert_tree_control(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        dialog_ctrl: XControl | None = None,
        **props: Any,
    ) -> CtlTree:
        """
        Create a new control of type Tree in the actual dialog.

        Args:
            x (int , UnitT): X coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            y (int , UnitT): Y coordinate in Pixels or ``UnitT``. If ``-1``, the dialog Position is not set.
            width (int , UnitT): Width in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            height (int , UnitT): Height in Pixels or ``UnitT``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            dialog_ctrl (XControl, optional): control. Defaults to class instance.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create Tree control.

        Returns:
            CtlTree: Tree Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        if dialog_ctrl is None:
            dialog_ctrl = self.__ctl
        with LoContext(inst=self.__lo_inst):
            result = self._DialogControlsPartial_dialogs_class.insert_tree_control(
                dialog_ctrl,
                x=x,
                y=y,
                width=width,
                height=height,
                border=border,
                name=name,
                **props,
            )
        return result

    # endregion Insert Tree Control

    @property
    def _DialogControlsPartial_dialogs_class(self) -> Type[Dialogs]:
        try:
            # avoid circular import.
            return self._dialogs_class_instance
        except AttributeError:
            # pylint: disable=import-outside-toplevel
            from ooodev.dialog.dialogs import Dialogs

            self._dialogs_class_instance = Dialogs
        return self._dialogs_class_instance

    # endregion Insert Controls


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dialogs import Dialogs
