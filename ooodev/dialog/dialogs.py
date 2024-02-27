# coding: utf-8
# pylint: disable=too-many-lines
# region Imports
from __future__ import annotations
import datetime
from typing import TYPE_CHECKING, Any, Iterable, List, Type, Tuple, cast, TypeVar
import uno


# pylint: disable=useless-import-alias
# pylint: disable=unused-import

from com.sun.star.awt import XControl
from com.sun.star.awt import XControlContainer
from com.sun.star.awt import XControlModel
from com.sun.star.awt import XDialog
from com.sun.star.awt import XDialogProvider
from com.sun.star.awt import XToolkit
from com.sun.star.awt import XTopWindow
from com.sun.star.awt import XWindow
from com.sun.star.awt.grid import XGridDataModel
from com.sun.star.awt.tree import XMutableTreeDataModel
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNameContainer
from com.sun.star.lang import XInitialization
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.lang import XServiceInfo

# com.sun.star.awt.PushButtonType
from ooo.dyn.awt.image_scale_mode import ImageScaleModeEnum as ImageScaleModeEnum
from ooo.dyn.awt.line_end_format import LineEndFormatEnum as LineEndFormatEnum
from ooo.dyn.awt.pos_size import PosSize as PosSize
from ooo.dyn.awt.push_button_type import PushButtonType as PushButtonType
from ooo.dyn.style.vertical_alignment import VerticalAlignment as VerticalAlignment
from ooo.dyn.view.selection_type import SelectionType as SelectionType
from ooo.dyn.style.horizontal_alignment import HorizontalAlignment as HorizontalAlignment
from ooo.dyn.awt.pos_size import PosSizeEnum as PosSizeEnum

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.kind.align_kind import AlignKind as AlignKind
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.date_format_kind import DateFormatKind as DateFormatKind
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.utils.kind.horz_ver_kind import HorzVertKind as HorzVertKind
from ooodev.utils.kind.orientation_kind import OrientationKind as OrientationKind
from ooodev.utils.kind.state_kind import StateKind as StateKind
from ooodev.utils.kind.time_format_kind import TimeFormatKind as TimeFormatKind
from ooodev.utils.kind.tri_state_kind import TriStateKind as TriStateKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase
from ooodev.dialog.dl_control.ctl_button import CtlButton
from ooodev.dialog.dl_control.ctl_check_box import CtlCheckBox
from ooodev.dialog.dl_control.ctl_combo_box import CtlComboBox
from ooodev.dialog.dl_control.ctl_currency_field import CtlCurrencyField
from ooodev.dialog.dl_control.ctl_date_field import CtlDateField
from ooodev.dialog.dl_control.ctl_dialog import CtlDialog
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
    from com.sun.star.awt import UnoControlButton  # service
    from com.sun.star.awt import UnoControlButtonModel  # service
    from com.sun.star.awt import UnoControlCheckBox  # service
    from com.sun.star.awt import UnoControlCheckBoxModel  # service
    from com.sun.star.awt import UnoControlComboBox  # service
    from com.sun.star.awt import UnoControlComboBoxModel  # service
    from com.sun.star.awt import UnoControlCurrencyField  # service
    from com.sun.star.awt import UnoControlCurrencyFieldModel  # service
    from com.sun.star.awt import UnoControlDateField
    from com.sun.star.awt import UnoControlDateFieldModel
    from com.sun.star.awt import UnoControlDialog  # service
    from com.sun.star.awt import UnoControlDialogModel  # service
    from com.sun.star.awt import UnoControlEdit  # service
    from com.sun.star.awt import UnoControlEditModel  # service
    from com.sun.star.awt import UnoControlFileControl  # service
    from com.sun.star.awt import UnoControlFileControlModel  # service
    from com.sun.star.awt import UnoControlFixedHyperlink  # service
    from com.sun.star.awt import UnoControlFixedHyperlinkModel  # service
    from com.sun.star.awt import UnoControlFixedLine  # service
    from com.sun.star.awt import UnoControlFixedLineModel  # service
    from com.sun.star.awt import UnoControlFixedText  # service
    from com.sun.star.awt import UnoControlFixedTextModel  # service
    from com.sun.star.awt import UnoControlFormattedField  # service
    from com.sun.star.awt import UnoControlFormattedFieldModel  # service
    from com.sun.star.awt import UnoControlGroupBox  # service
    from com.sun.star.awt import UnoControlGroupBoxModel  # service
    from com.sun.star.awt import UnoControlImageControl  # service
    from com.sun.star.awt import UnoControlImageControlModel  # service
    from com.sun.star.awt import UnoControlListBox  # service
    from com.sun.star.awt import UnoControlListBoxModel  # service
    from com.sun.star.awt import UnoControlNumericField  # service
    from com.sun.star.awt import UnoControlNumericFieldModel  # service
    from com.sun.star.awt import UnoControlPatternField  # service
    from com.sun.star.awt import UnoControlPatternFieldModel  # service
    from com.sun.star.awt import UnoControlProgressBar  # service
    from com.sun.star.awt import UnoControlProgressBarModel  # service
    from com.sun.star.awt import UnoControlRadioButton  # service
    from com.sun.star.awt import UnoControlRadioButtonModel  # service
    from com.sun.star.awt import UnoControlScrollBar  # service
    from com.sun.star.awt import UnoControlScrollBarModel  # service
    from com.sun.star.awt import UnoControlSpinButton  # service
    from com.sun.star.awt import UnoControlSpinButtonModel  # service
    from com.sun.star.awt import UnoControlTimeField  # service
    from com.sun.star.awt import UnoControlTimeFieldModel  # service
    from com.sun.star.awt.grid import UnoControlGrid  # service
    from com.sun.star.awt.grid import UnoControlGridModel  # service
    from com.sun.star.awt.tab import UnoControlTabPage  # service
    from com.sun.star.awt.tab import UnoControlTabPageContainer  # service
    from com.sun.star.awt.tab import UnoControlTabPageContainerModel  # service
    from com.sun.star.awt.tab import UnoControlTabPageModel  # service
    from com.sun.star.awt.tree import TreeControl  # service
    from com.sun.star.awt.tree import TreeControlModel  # service
    from com.sun.star.container import XNameAccess
    from com.sun.star.lang import EventObject
    from ooodev.utils.type_var import PathOrStr
# endregion Imports

if TYPE_CHECKING:
    ControlT = TypeVar("ControlT", bound=DialogControlBase)
else:
    ControlT = Any


class Dialogs:
    """Manages creating, accessing and inserting controls into dialogs"""

    StateEnum = TriStateKind

    # region    load & execute a dialog
    @staticmethod
    def load_dialog(script_name: str) -> XDialog:
        """
        Create a dialog for the given script name.

        |lo_unsafe|

        Args:
            script_name (str): script name.

        Raises:
            Exception: if unable to create dialog.

        Returns:
            XDialog: Dialog instance.
        """
        dp = mLo.Lo.create_instance_mcf(XDialogProvider, "com.sun.star.awt.DialogProvider", raise_err=True)

        try:
            return dp.createDialog(f"vnd.sun.star.script:{script_name}?location=application")
        except Exception as e:
            mLo.Lo.print("Could not access the Dialog Provider")
            raise e

    @staticmethod
    def load_addon_dialog(extension_id: str, dialog_fnm: str) -> XDialog:
        """
        Loads addon dialog.

        |lo_unsafe|

        Args:
            extension_id (str): Addon id.
            dialog_fnm (str): Addon file path.

        Raises:
            Exception: if unable to create dialog.

        Returns:
            XDialog: Dialog instance
        """
        dp = mLo.Lo.create_instance_mcf(XDialogProvider, "com.sun.star.awt.DialogProvider", raise_err=True)
        return dp.createDialog(f"vnd.sun.star.extension://{extension_id}/{dialog_fnm}")

    # endregion load & execute a dialog

    # region    access a control/component inside a dialog

    @staticmethod
    def find_control(dialog_ctrl: XControl, name: str) -> XControl | None:
        """
        Finds control by name.

        |lo_safe|

        Args:
            dialog_ctrl (XControl): Control.
            name (str): Name to find.

        Returns:
            XControl: Control if found, else ``None``.
        """
        ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
        return ctrl_con.getControl(name)

    @classmethod
    def find_controls(cls, dialog_ctrl: XControl, control_type: Type[ControlT]) -> List[ControlT]:
        """
        Finds controls by type.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            control_type (ControlT): Control type.

        Returns:
            List[ControlT]: List of controls.
        """
        ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
        controls = ctrl_con.getControls()
        result: List[ControlT] = []
        if not controls:
            return result
        for ctrl in controls:
            control = cls.get_dialog_control_instance(ctrl)
            if control and mInfo.Info.is_instance(control, control_type):
                result.append(control)
        return result

    @staticmethod
    def get_dialog_control_instance(
        dialog_ctrl: XControl,
    ) -> DialogControlBase | None:
        """
        Gets a control as a ``DialogControlBase`` control.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.

        Returns:
            DialogControlBase | None: Returns a ``DialogControlBase`` such as ``CtlButton`` or ``CtlCheckBox`` if found, else ``None``
        """
        si = mLo.Lo.qi(XServiceInfo, dialog_ctrl)
        if not si:
            return None
        named = DialogControlNamedKind.from_value(si.getImplementationName())
        if named == DialogControlNamedKind.UNKNOWN:
            return None
        if named == DialogControlNamedKind.BUTTON:
            return CtlButton(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.CHECKBOX:
            return CtlCheckBox(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.COMBOBOX:
            return CtlComboBox(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.CURRENCY:
            return CtlCurrencyField(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.DATE_FIELD:
            return CtlDateField(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.FILE_CONTROL:
            return CtlFile(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.FIXED_LINE:
            return CtlFixedLine(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.FIXED_TEXT:
            return CtlFixedText(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.FORMATTED_TEXT:
            return CtlFormattedField(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.GRID_CONTROL:
            return CtlGrid(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.GROUP_BOX:
            return CtlGroupBox(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.HYPERLINK:
            return CtlHyperlinkFixed(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.IMAGE:
            return CtlImage(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.LIST_BOX:
            return CtlListBox(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.NUMERIC:
            return CtlNumericField(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.PATTERN:
            return CtlPatternField(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.PROGRESS_BAR:
            return CtlProgressBar(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.RADIO_BUTTON:
            return CtlRadioButton(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.SCROLL_BAR:
            return CtlScrollBar(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.SPIN_BUTTON:
            return CtlSpinButton(dialog_ctrl)
        if named == DialogControlNamedKind.TAB_PAGE_CONTAINER:
            return CtlTabPageContainer(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.TAB_PAGE:
            return CtlTabPage(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.EDIT:
            return CtlTextEdit(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.TIME:
            return CtlTimeField(dialog_ctrl)  # type: ignore
        if named == DialogControlNamedKind.TREE:
            return CtlTree(dialog_ctrl)  # type: ignore
        return None

    @classmethod
    def show_control_info(cls, dialog_ctrl: XControl) -> None:
        """
        Prints info for a control to console.

        |lo_safe|

        Args:
            dialog_ctrl (XControl): Control

        Returns:
            None:
        """
        controls = cls.get_dialog_controls_arr(dialog_ctrl)
        print(f"No of controls: {len(controls)}")
        for i, ctl in enumerate(controls):
            print(f"{i}. Name: {cls.get_control_name(ctl)}")
            print(f"  Default Control: {cls.get_control_class_id(ctl)}")
            print()

    @staticmethod
    def get_dialog_controls_arr(dialog_ctrl: XControl) -> Tuple[XControl, ...]:
        """
        Gets all controls for a given control.

        |lo_safe|

        Args:
            dialog_ctrl (XControl): control.

        Returns:
            Tuple[XControl, ...]: controls.
        """
        ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
        return ctrl_con.getControls()

    # @staticmethod
    # def get_controls(dialog_ctrl: XControl, control_name: str = "") -> Tuple[str, ...]:
    #     na = mLo.Lo.qi(XNameAccess, dialog_ctrl.getModel(), True)
    #     element_names = na.getElementNames()
    #     if not control_name:
    #         return element_names
    #     if na.hasByName(control_name):
    #         return [na.getByName(control_name)]

    @staticmethod
    def get_control_props(control_model: Any) -> XPropertySet:
        """
        Gets property set for a control model.

        |lo_safe|

        Args:
            control_model (Any): control model.

        Returns:
            XPropertySet: Property set.
        """
        return mLo.Lo.qi(XPropertySet, control_model, True)

    @classmethod
    def get_control_name(cls, control: XControl) -> str:
        """
        Get the name of a control.

        |lo_safe|

        Args:
            control (XControl): control.

        Returns:
            str: control name.
        """
        props = cls.get_control_props(control.getModel())
        return str(props.getPropertyValue("Name"))

    @classmethod
    def get_control_class_id(cls, control: XControl) -> str:
        """
        Gets control class id.

        |lo_safe|

        Args:
            control (XControl): control.

        Returns:
            str: class id.
        """
        props = cls.get_control_props(control.getModel())
        return str(props.getPropertyValue("DefaultControl"))

    @classmethod
    def get_event_source_name(cls, event: EventObject) -> str:
        """
        Get event source name.

        |lo_safe|

        Args:
            event (EventObject): event.

        Returns:
            str: event source name.
        """
        return cls.get_control_name(cls.get_event_control(event))

    @staticmethod
    def get_event_control(event: EventObject) -> XControl:
        """
        Gets event control from event.

        |lo_safe|

        Args:
            event (EventObject): event.

        Returns:
            XControl: control.
        """
        return mLo.Lo.qi(XControl, event.Source, True)

    # endregion access a control/component inside a dialog

    # region    convert dialog into other forms
    @staticmethod
    def get_dialog(dialog_ctrl: XControl) -> XDialog:
        """
        Gets dialog from dialog control.

        |lo_safe|

        Args:
            dialog_ctrl (XControl): control.

        Returns:
            XDialog: dialog.
        """
        return mLo.Lo.qi(XDialog, dialog_ctrl, True)

    @staticmethod
    def get_dialog_control(dialog: XDialog) -> XControl:
        """
        Gets dialog control.

        |lo_safe|

        Args:
            dialog (XDialog): dialog.

        Returns:
            XControl: control.
        """
        return mLo.Lo.qi(XControl, dialog, True)

    @staticmethod
    def get_dialog_window(dialog_ctrl: XControl) -> XTopWindow:
        """
        Gets dialog window.

        |lo_safe|

        Args:
            dialog_ctrl (XControl): dialog control.

        Returns:
            XTopWindow: Top window instance.
        """
        return mLo.Lo.qi(XTopWindow, dialog_ctrl, True)

    # endregion convert dialog into other forms

    # region    create a dialog
    @classmethod
    def create_dialog(
        cls,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        title: str,
        **props: Any,
    ) -> CtlDialog:
        """
        Creates a dialog.

        |lo_unsafe|

        Args:
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int): Height. If ``-1``, the dialog Size is not set.
            title (str): Dialog title.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create dialog.

        Returns:
            CtlDialog: Control.

        See Also:
            `API UnoControlDialogModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogModel.html>`_
        """
        # sourcery skip: raise-specific-error
        try:
            dialog_ctrl = cast(
                "UnoControlDialog",
                mLo.Lo.create_instance_mcf(XControl, "com.sun.star.awt.UnoControlDialog", raise_err=True),
            )
            control_model = mLo.Lo.create_instance_mcf(
                XControlModel, "com.sun.star.awt.UnoControlDialogModel", raise_err=True
            )
            dialog_ctrl.setModel(control_model)

            ctl_props = cls.get_control_props(dialog_ctrl.getModel())

            ctl_props.setPropertyValue("Title", title)
            ctl_props.setPropertyValue("Name", "OfficeDialog")

            ctl_props.setPropertyValue("Step", 0)
            ctl_props.setPropertyValue("Moveable", True)
            ctl_props.setPropertyValue("TabIndex", 0)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            cls._set_size_pos(dialog_ctrl, x, y, width, height)
            return CtlDialog(dialog_ctrl)
        except Exception as e:
            raise mEx.DialogError(f"Could not create dialog control: {e}") from e

    @classmethod
    def create_dialog_peer(cls, dialog_ctrl: XControl | CtlDialog) -> XDialog:
        """
        Creates a dialog peer.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.

        Returns:
            XDialog: Dialog.
        """
        if mInfo.Info.is_instance(dialog_ctrl, CtlDialog):
            ctrl = dialog_ctrl.control
        else:
            ctrl = cast("UnoControlDialog", dialog_ctrl)
        xwindow = mLo.Lo.qi(XWindow, ctrl, True)
        # set the dialog window invisible until it is executed
        xwindow.setVisible(False)

        xtoolkit = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit", raise_err=True)
        window_parent_peer = xtoolkit.getDesktopWindow()
        ctrl.createPeer(xtoolkit, window_parent_peer)

        # dialog_component = mLo.Lo.qi(XComponent, dialog_ctrl)
        # dialog_component.dispose() # free window resources
        # commented out or the Add-on dialog crashes when called a second time
        # because createPeer() cannot find a model
        return cls.get_dialog(ctrl)

    # endregion create a dialog

    # region    add components to a dialog

    @staticmethod
    def get_dialog_nm_con(ctrl: XControl) -> XNameContainer:
        """
        Gets Name container from control.

        |lo_safe|

        Args:
            ctrl (XControl): Dialog control.

        Returns:
            XNameContainer: Name Container.
        """
        return mLo.Lo.qi(XNameContainer, ctrl.getModel(), True)

    @staticmethod
    def create_name(elem_container: XNameAccess, name: str) -> str:
        """
        Creates a name.

        Make a unique string by appending a number to the supplied name.

        |lo_safe|

        Args:
            elem_container (XNameAccess): container.
            name (str): current name.

        Returns:
            str: a name not in container.
        """
        used_name = True
        i = 1
        nm = f"{name}{i}"
        while used_name:
            used_name = elem_container.hasByName(nm)
            if used_name:
                i += 1
                nm = f"{name}{i}"
        return nm

    @classmethod
    def insert_button(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        btn_type: PushButtonType | None = None,
        name: str = "",
        **props: Any,
    ) -> CtlButton:
        """
        Insert Button Control.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            label (str): Button Label.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            btn_type (PushButtonType | None, optional): Type of Button.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
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
        # sourcery skip: raise-specific-error
        if btn_type is None:
            btn_type = PushButtonType.STANDARD
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlButtonModel", msf.createInstance("com.sun.star.awt.UnoControlButtonModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "CommandButton")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Label", label)
            # ctl_props.setPropertyValue("PushButtonType", btn_type)
            uno_any = uno.Any("short", btn_type)  # type: ignore
            uno.invoke(ctl_props, "setPropertyValue", ("PushButtonType", uno_any))  # type: ignore
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlButton", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlButton(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create button control: {e}") from e

    @classmethod
    def insert_check_box(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 8,
        tri_state: bool = True,
        state: TriStateKind = TriStateKind.CHECKED,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlCheckBox:
        """
        Inserts a check box control.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            label (str): Checkbox label text.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``8``. If ``-1``, the dialog Size is not set.
            tri_state (TriStateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``TriStateKind.CHECKED``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of Checkbox. Must be a unique name. If empty, a unique name is generated.
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
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlCheckBoxModel", msf.createInstance("com.sun.star.awt.UnoControlCheckBoxModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "CheckBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            # Checkboxes do not have a border property but does have a VisualEffect property that is basically the same
            ctl_props.setPropertyValue("VisualEffect", int(border))
            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("Label", label)
            ctl_props.setPropertyValue("TriState", tri_state)
            ctl_props.setPropertyValue("State", int(state))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlCheckBox", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlCheckBox(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create check box control: {e}") from e

    @classmethod
    def insert_combo_box(
        cls,
        dialog_ctrl: XControl,
        *,
        entries: Iterable[str],
        x: int,
        y: int,
        width: int,
        height: int = 20,
        max_text_len: int = 0,
        drop_down: bool = True,
        read_only: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlComboBox:
        """
        Insert a combo box control.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            entries (Iterable[str]): Combo box entries.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            max_text_len (int, optional): Specifies the maximum character count, There's no limitation, if set to 0. Defaults to ``0``.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``True``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
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
        # sourcery skip: raise-specific-error
        try:
            max_text_len = max(max_text_len, 0)
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlComboBoxModel", msf.createInstance("com.sun.star.awt.UnoControlComboBoxModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ComboBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("Dropdown", drop_down)
            ctl_props.setPropertyValue("MaxTextLen", max_text_len)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("ReadOnly", read_only)
            if entries:
                uno_strings = uno.Any("[]string", tuple(entries))  # type: ignore
                uno.invoke(ctl_props, "setPropertyValue", ("StringItemList", uno_strings))  # type: ignore
                # ctl_props.setPropertyValue("StringItemList", tuple(entries))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlComboBox", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlComboBox(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create combo box control: {e}") from e

    @classmethod
    def insert_currency_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        value: float = 0.0,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlCurrencyField:
        """
        Inserts a currency control.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            value (float, optional): Control Value. Defaults to ``0.0``.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create currency field box control.

        Returns:
            CtlCurrencyField: Currency field control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlCurrencyFieldModel", msf.createInstance("com.sun.star.awt.UnoControlCurrencyFieldModel")
            )
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "CurrencyField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("ValueMin", min_value)
            ctl_props.setPropertyValue("ValueMax", max_value)
            ctl_props.setPropertyValue("ValueStep", increment)
            ctl_props.setPropertyValue("Spin", spin_button)
            ctl_props.setPropertyValue("Value", value)
            ctl_props.setPropertyValue("DecimalAccuracy", accuracy)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlCurrencyField", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlCurrencyField(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create currency field control: {e}") from e

    @classmethod
    def insert_date_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        date_value: datetime.datetime | None = None,
        min_date: datetime.datetime = datetime.datetime(1900, 1, 1, 0, 0, 0, 0),
        max_date: datetime.datetime = datetime.datetime(2200, 12, 31, 0, 0, 0, 0),
        drop_down: bool = True,
        date_format: DateFormatKind = DateFormatKind.SYSTEM_SHORT,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlDateField:
        """
        Create a new control of type DateField in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            date_value (datetime.datetime | None, optional): Specifics control datetime. Defaults to ``None``.
            min_date (datetime.datetime, optional): Specifics control min datetime. Defaults to ``datetime(1900, 1, 1, 0, 0, 0, 0)``.
            max_date (datetime.datetime, optional): Specifics control Min datetime. Defaults to ``datetime(2200, 12, 31, 0, 0, 0, 0)``.
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to ``True``.
            date_format (DateFormatKind, optional): Date format. Defaults to ``DateFormatKind.SYSTEM_SHORT``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create date field control.

        Returns:
            CtlDateField: Date field control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``DateFormatKind`` can be imported from ``ooodev.utils.kind.date_format_kind``.
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlDateFieldModel", msf.createInstance("com.sun.star.awt.UnoControlDateFieldModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "DateField")

            # set properties in the model
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Dropdown", drop_down)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("DateMin", DateUtil.date_to_uno_date(min_date))
            ctl_props.setPropertyValue("DateMax", DateUtil.date_to_uno_date(max_date))
            if date_value is not None:
                ctl_props.setPropertyValue("Date", DateUtil.date_to_uno_date(date_value))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlDateField", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            ctl = CtlDateField(result)
            ctl.date_format = date_format
            return ctl
        except Exception as e:
            raise mEx.DialogError(f"Could not create text field control: {e}") from e

    @classmethod
    def insert_file_control(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlFile:
        """
        Create a new control of type FileControl in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create file control.

        Returns:
            CtlFile: File Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlFileControlModel", msf.createInstance("com.sun.star.awt.UnoControlFileControlModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "FileControl")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlFileControl", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlFile(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create file control: {e}") from e

    @classmethod
    def insert_fixed_line(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 1,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        name: str = "",
        **props: Any,
    ) -> CtlFixedLine:
        """
        Create a new control of type Fixed Line in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``1``. If ``-1``, the dialog Size is not set.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create fixed line control.

        Returns:
            CtlFixedLine: Fixed Line Control.

        Hint:
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlFixedLineModel", msf.createInstance("com.sun.star.awt.UnoControlFixedLineModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "FixedLine")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Orientation", int(orientation))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlFixedLine", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlFixedLine(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create fixed line control: {e}") from e

    @classmethod
    def insert_formatted_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        value: float | None = None,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlFormattedField:
        """
        Create a new control of type FormattedField in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            value (float, optional): Control Value. Defaults to ``0.0``.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create formatted field control.

        Returns:
            CtlFormattedField: Formatted Field.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlFormattedFieldModel", msf.createInstance("com.sun.star.awt.UnoControlFormattedFieldModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "FormattedField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("EffectiveMin", min_value)
            ctl_props.setPropertyValue("EffectiveMax", max_value)
            ctl_props.setPropertyValue("Spin", spin_button)
            ctl_props.setPropertyValue("Border", int(border))
            if value is not None:
                ctl_props.setPropertyValue("EffectiveValue", value)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlFormattedField", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlFormattedField(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create formatted field control: {e}") from e

    @classmethod
    def insert_group_box(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        label: str = "",
        name: str = "",
        **props: Any,
    ) -> CtlGroupBox:
        """
        Create a new control of type GroupBox in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. If ``-1``, the dialog Size is not set.
            label (str, optional): Group box label.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create group box control.

        Returns:
            CtlGroupBox: Group box Control.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlGroupBoxModel", msf.createInstance("com.sun.star.awt.UnoControlGroupBoxModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "GroupBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Name", name)
            if label:
                ctl_props.setPropertyValue("Label", label)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlGroupBox", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlGroupBox(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create Group box control: {e}") from e

    @classmethod
    def insert_hyperlink(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        url: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        align: AlignKind = AlignKind.LEFT,
        vert_align: VerticalAlignment = VerticalAlignment.TOP,
        multiline: bool = False,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        **props: Any,
    ) -> CtlHyperlinkFixed:
        """
        Create a new control of type Hyperlink in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            label (str): Hyperlink label.
            url (str): Hyperlink URL.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            align (AlignKind, optional): Horizontal alignment. Defaults to ``AlignKind.LEFT``.
            vert_align (VerticalAlignment, optional): Vertical alignment. Defaults to ``VerticalAlignment.TOP``.
            multiline (bool, optional): Specifies if the control can display multiple lines of text. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
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
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlFixedHyperlinkModel", msf.createInstance("com.sun.star.awt.UnoControlFixedHyperlinkModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "Hyperlink")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Align", align.value)
            ctl_props.setPropertyValue("VerticalAlign", vert_align)
            ctl_props.setPropertyValue("MultiLine", multiline)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)

            if label:
                ctl_props.setPropertyValue("Label", label)
            if url:
                ctl_props.setPropertyValue("URL", url)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlFixedHyperlink", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlHyperlinkFixed(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create Hyperlink control: {e}") from e

    @classmethod
    def insert_image_control(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        border: BorderKind = BorderKind.BORDER_3D,
        scale: int | ImageScaleModeEnum = ImageScaleModeEnum.NONE,
        image_url: PathOrStr = "",
        name: str = "",
        **props: Any,
    ) -> CtlImage:
        """
        Create a new control of type ImageControl in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            image_url (PathOrStr, optional): Image URL. When setting the value it can be a string or a Path object.
                If a string is passed it can be a URL or a path to a file.
                Value such as ``file:///path/to/image.png`` and ``/path/to/image.png`` are valid.
                Relative paths are supported.
            scale (int | ImageScaleModeEnum, optional): Image scale mode. Defaults to ``ImageScaleModeEnum.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
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
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlImageControlModel", msf.createInstance("com.sun.star.awt.UnoControlImageControlModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ImageControl")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)

            scale = int(scale)
            if scale != ImageScaleModeEnum.NONE.value:
                ctl_props.setPropertyValue("ScaleImage", True)
                ctl_props.setPropertyValue("ScaleMode", scale)
            else:
                ctl_props.setPropertyValue("ScaleImage", False)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlImageControl", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            ctl_image = CtlImage(result)
            if image_url:
                ctl_image.picture = image_url
            return ctl_image

        except Exception as e:
            raise mEx.DialogError(f"Could not create image control: {e}") from e

    @classmethod
    def insert_label(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        name: str = "",
        **props: Any,
    ) -> CtlFixedText:
        """
        Insert a Fixed Text into a control.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            label (str): Contents of label.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Default ``20``. If ``-1``, the dialog Size is not set.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create Fixed Text.

        Returns:
            CtlFixedText: Fixed Text Control.

        See Also:
            `API UnoControlFixedTextModel Service <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlFixedTextModel.html>`_
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)

            model = cast("UnoControlFixedTextModel", msf.createInstance("com.sun.star.awt.UnoControlFixedTextModel"))

            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "FixedText")

            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Label", label)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
            result = cast("UnoControlFixedText", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlFixedText(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create Fixed Text control: {e}") from e

    @classmethod
    def insert_list_box(
        cls,
        dialog_ctrl: XControl,
        *,
        entries: Iterable[str],
        x: int,
        y: int,
        width: int,
        height: int = 100,
        drop_down: bool = False,
        read_only: bool = False,
        line_count: int = 5,
        multi_select: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlListBox:
        """
        Insert a list box control.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            entries (Iterable[str]): List box entries.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``100``. If ``-1``, the dialog Size is not set.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``False``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            line_count (int, optional): Specifies the number of lines to display. Defaults to ``5``.
            multi_select (int, optional): Specifies if multiple entries can be selected. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create list box control.

        Returns:
            CtlListBox: List box control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlListBoxModel", msf.createInstance("com.sun.star.awt.UnoControlListBoxModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ListBox")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("LineCount", line_count)
            ctl_props.setPropertyValue("Dropdown", drop_down)
            ctl_props.setPropertyValue("MultiSelection", multi_select)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("ReadOnly", read_only)
            # if entries:
            #     uno_strings = uno.Any("[]string", tuple(entries))  # type: ignore
            #     uno.invoke(ctl_props, "setPropertyValue", ("StringItemList", uno_strings))  # type: ignore

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlListBox", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            ctl = CtlListBox(result)
            ctl.set_list_data(entries)
            return ctl
        except Exception as e:
            raise mEx.DialogError(f"Could not create list box control: {e}") from e

    @classmethod
    def insert_password_field(
        cls,
        dialog_ctrl: XControl,
        *,
        text: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        **props: Any,
    ) -> CtlTextEdit:
        """
        Inserts a password field.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            text (str): Text value.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create text field.

        Returns:
            CtlTextEdit: Text Field Control.

        See Also:
            :py:meth:`~.dialogs.Dialogs.insert_text_field`.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        return cls.insert_text_field(
            dialog_ctrl=dialog_ctrl,
            text=text,
            x=x,
            y=y,
            width=width,
            height=height,
            echo_char="*",
            border=border,
            name=name,
            **props,
        )

    @classmethod
    def insert_pattern_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        edit_mask: str = "",
        literal_mask: str = "",
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlPatternField:
        """
        Create a new control of type PatternField in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Default``20``. If ``-1``, the dialog Size is not set.
            edit_mask (str, optional): Specifies a character code that determines what the user may enter. Defaults to ``""``.
            literal_mask (str, optional): Specifies the initial values that are displayed in the pattern field. Defaults to ``""``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create pattern field control.

        Returns:
            CtlPatternField: Pattern Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlPatternFieldModel", msf.createInstance("com.sun.star.awt.UnoControlPatternFieldModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "PatternField")
            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("EditMask", edit_mask)
            ctl_props.setPropertyValue("LiteralMask", literal_mask)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlPatternField", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlPatternField(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create pattern field control: {e}") from e

    @classmethod
    def insert_numeric_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        value: float | None = None,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlNumericField:
        """
        Create a new control of type NumericField in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Default ``20``. If ``-1``, the dialog Size is not set.
            value (float, optional): Control Value. Defaults to ``0.0``.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create numeric field control.

        Returns:
            CtlNumericField: Numeric Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlNumericFieldModel", msf.createInstance("com.sun.star.awt.UnoControlNumericFieldModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "NumericField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("ValueMin", min_value)
            ctl_props.setPropertyValue("ValueMax", max_value)
            ctl_props.setPropertyValue("ValueStep", increment)
            ctl_props.setPropertyValue("DecimalAccuracy", accuracy)
            ctl_props.setPropertyValue("Spin", spin_button)
            ctl_props.setPropertyValue("Name", name)
            if value is not None:
                ctl_props.setPropertyValue("Value", value)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlNumericField", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlNumericField(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create numeric field control: {e}") from e

    @classmethod
    def insert_progress_bar(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        min_value: int = 0,
        max_value: int = 100,
        value: int = 0,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlProgressBar:
        """
        Create a new control of type Progress Control in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int): Height. If ``-1``, the dialog Size is not set.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            value (int, optional): The value initial value of the progress bar. Defaults to ``0``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create numeric field control.

        Returns:
            CtlProgressBar: Progress Bar Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast(
                "UnoControlProgressBarModel", msf.createInstance("com.sun.star.awt.UnoControlProgressBarModel")
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ProgressBar")
            # set properties in the model
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("ProgressValueMin", min_value)
            ctl_props.setPropertyValue("ProgressValueMax", max_value)
            ctl_props.setPropertyValue("ProgressValue", value)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlProgressBar", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlProgressBar(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create Progress Bar control: {e}") from e

    @classmethod
    def insert_radio_button(
        cls,
        dialog_ctrl: XControl,
        *,
        label: str,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        multiline: bool = False,
        name: str = "",
        **props: Any,
    ) -> CtlRadioButton:
        """
        Create a new control of type RadioButton in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            label (str): Contents of label.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Default ``20``. If ``-1``, the dialog Size is not set.
            multiline (bool, optional): Specifies if the control can display multiple lines of text. Defaults to ``False``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create radio button.

        Returns:
            CtlRadioButton: Radio button control.
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)

            model = cast(
                "UnoControlRadioButtonModel", msf.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
            )

            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "RadioButton")

            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Label", label)
            ctl_props.setPropertyValue("MultiLine", multiline)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)
            result = cast("UnoControlRadioButton", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlRadioButton(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create radio button control: {e}") from e

    @classmethod
    def insert_scroll_bar(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        min_value: int = 0,
        max_value: int = 100,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlScrollBar:
        """
        Create a new control of type ScrollBar in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int): Height. If ``-1``, the dialog Size is not set.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create scroll bar control.

        Returns:
            CtlScrollBar: Scroll Bar Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlScrollBarModel", msf.createInstance("com.sun.star.awt.UnoControlScrollBarModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "ScrollBar")
            # set properties in the model
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Orientation", int(orientation))
            ctl_props.setPropertyValue("ScrollValueMin", min_value)
            ctl_props.setPropertyValue("ScrollValueMax", max_value)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlScrollBar", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlScrollBar(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create scroll bar control: {e}") from e

    @classmethod
    def insert_spin_button(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        min_value: int = 0,
        max_value: int = 100,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlSpinButton:
        """
        Create a new control of type Spin Button in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int): Height. If ``-1``, the dialog Size is not set.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create spin button control.

        Returns:
            CtlSpinButton: Spin Button Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.

        .. versionadded:: 0.29.0
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlSpinButtonModel", msf.createInstance("com.sun.star.awt.UnoControlSpinButtonModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "SpinButton")
            # set properties in the model
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Orientation", int(orientation))
            ctl_props.setPropertyValue("SpinValueMax", max_value)
            ctl_props.setPropertyValue("SpinValueMin", min_value)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlSpinButton", ctrl_con.getControl(name))
            btn = CtlSpinButton(result)
            # not sure why but this control seems buggy with setting size and position.
            # Setting Width and height seems to have the best results when setting the model width and height.

            if width > -1 and height > -1:
                btn.model.Width = width
                btn.model.Height = height

            # only set position here. Width and height set above.
            cls._set_size_pos(result, x, y, -1, -1)
            # cls._set_size_pos(result, x, y, width, height)
            return btn

        except Exception as e:
            raise mEx.DialogError(f"Could not create scroll bar control: {e}") from e

    @classmethod
    def insert_tab_control(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 1,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
    ) -> CtlTabPageContainer:
        """
        Create a new control of type tab in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``1``. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.

        Raises:
            DialogError: If unable to create tab control.

        Returns:
            CtlTabPageContainer: Tab Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.

        See Also:
            :py:meth:`~.dialogs.Dialogs.insert_tab_page`.
        """
        try:
            dialog = cast("UnoControlDialog", cls.get_dialog(dialog_ctrl))
            if not dialog or not cast(XServiceInfo, dialog).supportsService("com.sun.star.awt.UnoControlDialog"):
                raise mEx.DialogError("Could not get dialog")
            dialog_model = cast("UnoControlDialogModel", dialog.getModel())
            model = cast(
                "UnoControlTabPageContainerModel",
                dialog_model.createInstance("com.sun.star.awt.tab.UnoControlTabPageContainerModel"),
            )

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog)
            if name:
                model.Name = name
            else:
                model.Name = cls.create_name(name_con, "TabControl")
            if border != BorderKind.NONE and hasattr(model, "Border"):
                setattr(model, "Border", int(border))
            # Add the model to the dialog
            dialog_model.insertByName(model.Name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlTabPageContainer", ctrl_con.getControl(model.Name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlTabPageContainer(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create Tab control: {e}") from e

    @classmethod
    def insert_tab_page(
        cls,
        dialog_ctrl: XControl,
        *,
        tab_ctrl: CtlTabPageContainer,
        title: str,
        tab_position: int,
        name: str = "",
        **props: Any,
    ) -> CtlTabPage:
        """
        Create a new control of type Tab Page in the actual tab control.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            tab_ctrl (CtlTabPageContainer): Tab Container.
            title (str): Tab title.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create tab page control.

        Returns:
            CtlTabPage: Tab Page Control.

        See Also:
            :py:meth:`~.dialogs.Dialogs.insert_tab_control`.
        """

        def create_name(ctl: UnoControlTabPageContainer, name: str) -> str:
            items = cast("Tuple[UnoControlTabPage, ...]", ctl.Controls)  # type: ignore
            i = 1
            nm = f"{name}{i}"
            for itm in items:
                if not cast(XServiceInfo, itm).supportsService("com.sun.star.awt.tab.UnoControlTabPage"):
                    continue
                itm_model = cast("UnoControlTabPageModel", itm.getModel())
                ctl_props = cls.get_control_props(itm_model)
                prop_name = ctl_props.getPropertyValue("Name")
                if prop_name and prop_name == nm:
                    i += 1
                    nm = f"{name}{i}"
                else:
                    break
            return nm

        try:
            if not tab_ctrl or not cast(XServiceInfo, tab_ctrl.view).supportsService(
                "com.sun.star.awt.tab.UnoControlTabPageContainer"
            ):
                raise mEx.DialogError("Not a valid UnoControlTabPageContainer")
            dialog = cast("UnoControlDialog", cls.get_dialog(dialog_ctrl))
            if not dialog or not cast(XServiceInfo, dialog).supportsService("com.sun.star.awt.UnoControlDialog"):
                raise mEx.DialogError("Could not get dialog")
            dialog_model = cast("UnoControlDialogModel", dialog.getModel())
            model = cast(
                "UnoControlTabPageModel", dialog_model.createInstance("com.sun.star.awt.tab.UnoControlTabPageModel")
            )

            tab_ctrl_model = tab_ctrl.model
            model_init = mLo.Lo.qi(XInitialization, model, True)
            model_init.initialize((tab_position,))

            model.Title = title
            if not name:
                name = create_name(tab_ctrl.view, "TabPage")

            # set properties in the model
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            index = tab_position - 1
            tab_ctrl_model.insertByIndex(index, model)
            controls = cast("Tuple[UnoControlTabPage, ...]", tab_ctrl.view.Controls)  # type: ignore
            if len(controls) > index:
                return CtlTabPage(controls[index])
            return None  # type: ignore

        except Exception as e:
            raise mEx.DialogError(f"Could not create Tab Page control: {e}") from e

    @classmethod
    def insert_table_control(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        row_header: bool = True,
        col_header: bool = True,
        grid_lines: bool = False,
        scroll_bars: HorzVertKind = HorzVertKind.NONE,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlGrid:
        """
        Create a new control of type TableControl in the actual dialog.

        To add data to the table use :py:meth:`~.dialogs.Dialogs.set_table_data`.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int): Height. If ``-1``, the dialog Size is not set.
            row_header (bool, optional): Specifies if the control has a row header. Defaults to ``True``.
            col_header (bool, optional): Specifies if the control has a column header. Defaults to ``True``.
            grid_lines (bool, optional): Specifies if the control has grid lines. when True horizontal and vertical lines are painted between the grid cells. Defaults to ``False``.
            scroll_bars (HorzVertKind, optional): Specifies if the control has scroll bars. Scrollbars always appear dynamically when they are needed. Defaults to ``HorzVertKind.NONE``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create table control.

        Returns:
            CtlGrid: Table Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``HorzVertKind`` can be imported from ``ooodev.utils.kind.horz_ver_kind``.

        See Also:
            :py:meth:`~.dialogs.Dialogs.set_table_data`.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlGridModel", msf.createInstance("com.sun.star.awt.grid.UnoControlGridModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "Grid")

            # set properties in the model
            if hasattr(model, "Border"):
                setattr(model, "Border", int(border))

            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("HScroll", HorzVertKind.HORIZONTAL in scroll_bars)
            ctl_props.setPropertyValue("VScroll", HorzVertKind.VERTICAL in scroll_bars)
            ctl_props.setPropertyValue("ShowColumnHeader", col_header)
            ctl_props.setPropertyValue("ShowRowHeader", row_header)
            ctl_props.setPropertyValue("UseGridLines", grid_lines)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # set the data model
            model.GridDataModel = mLo.Lo.create_instance_mcf(
                XGridDataModel, "com.sun.star.awt.grid.DefaultGridDataModel", raise_err=True
            )

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlGrid", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlGrid(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create Table control: {e}") from e

    @classmethod
    def insert_text_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        text: str = "",
        echo_char: str = "",
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        **props: Any,
    ) -> CtlTextEdit:
        """
        Inserts a text Field.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            text (str, optional): Text value.
            echo_char (str, optional): Character used for masking. Must be a single character.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
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
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlEditModel", msf.createInstance("com.sun.star.awt.UnoControlEditModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "TextField")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            if text:
                ctl_props.setPropertyValue("Text", text)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            if len(echo_char) == 1:  # for password fields
                ctl_props.setPropertyValue("EchoChar", ord(echo_char))

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlEdit", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            return CtlTextEdit(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create text field control: {e}") from e

    @classmethod
    def insert_tree_control(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlTree:
        """
        Create a new control of type Tree in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int): Height. If ``-1``, the dialog Size is not set.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create Tree control.

        Returns:
            CtlTree: Tree Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("TreeControlModel", msf.createInstance("com.sun.star.awt.tree.TreeControlModel"))

            # generate a unique name for the control
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "Tree")

            # set properties in the model
            # inherited from UnoControlDialogElement and UnoControlButtonModel
            ctl_props = cls.get_control_props(model)
            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("SelectionType", SelectionType.SINGLE)
            ctl_props.setPropertyValue("Editable", False)
            ctl_props.setPropertyValue("ShowsHandles", True)
            ctl_props.setPropertyValue("ShowsRootHandles", True)

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # set the data model
            model.DataModel = mLo.Lo.create_instance_mcf(
                XMutableTreeDataModel, "com.sun.star.awt.tree.MutableTreeDataModel", raise_err=True
            )

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # get the dialog's container holding all the control views
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("TreeControl", ctrl_con.getControl(name))
            # TreeControl does implement XWindows event thought it is documented
            cls._set_size_pos(result, x, y, width, height)  # type: ignore
            return CtlTree(result)
        except Exception as e:
            raise mEx.DialogError(f"Could not create Tree control: {e}") from e

    @classmethod
    def insert_time_field(
        cls,
        dialog_ctrl: XControl,
        *,
        x: int,
        y: int,
        width: int,
        height: int = 20,
        time_value: datetime.time | None = None,
        min_time: datetime.time = datetime.time(0, 0, 0, 0),
        max_time: datetime.time = datetime.time(23, 59, 59, 999_999),
        time_format: TimeFormatKind = TimeFormatKind.SHORT_24H,
        spin_button: bool = True,
        border: BorderKind = BorderKind.BORDER_3D,
        name: str = "",
        **props: Any,
    ) -> CtlTimeField:
        """
        Create a new control of type TimeField in the actual dialog.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            x (int): X coordinate. If ``-1``, the dialog Position is not set.
            y (int): Y coordinate. If ``-1``, the dialog Position is not set.
            width (int): Width. If ``-1``, the dialog Size is not set.
            height (int, optional): Height. Defaults to ``20``. If ``-1``, the dialog Size is not set.
            time_value (datetime.time | None, optional): Specifics the control time. Defaults to ``None``.
            min_time (datetime.time, optional): Specifics control min time. Defaults to ``time(0, 0, 0, 0)``.
            max_time (datetime.time, optional): Specifics control min time. Defaults to a ``time(23, 59, 59, 999_999)``.
            time_format (TimeFormatKind, optional): Date format. Defaults to ``TimeFormatKind.SHORT_24H``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``True``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            name (str, optional): Name of button. Must be a unique name. If empty, a unique name is generated.
            props (dict, optional): Extra properties to set for control.

        Raises:
            DialogError: If unable to create time field control.

        Returns:
            CtlTimeField: Time field control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``TimeFormatKind`` can be imported from ``ooodev.utils.kind.time_format_kind``.
        """
        # sourcery skip: raise-specific-error
        try:
            msf = mLo.Lo.qi(XMultiServiceFactory, dialog_ctrl.getModel(), True)
            model = cast("UnoControlTimeFieldModel", msf.createInstance("com.sun.star.awt.UnoControlTimeFieldModel"))
            name_con = cls.get_dialog_nm_con(dialog_ctrl)
            if not name:
                name = cls.create_name(name_con, "TimeField")

            # set properties in the model
            ctl_props = cls.get_control_props(model)

            ctl_props.setPropertyValue("Border", int(border))
            ctl_props.setPropertyValue("Name", name)
            ctl_props.setPropertyValue("TimeMin", DateUtil.time_to_uno_time(min_time))
            ctl_props.setPropertyValue("TimeMax", DateUtil.time_to_uno_time(max_time))
            ctl_props.setPropertyValue("Spin", spin_button)
            if time_value is not None:
                ctl_props.setPropertyValue("Time", DateUtil.time_to_uno_time(time_value))

            # set any extra user properties
            for k, v in props.items():
                ctl_props.setPropertyValue(k, v)

            # Add the model to the dialog
            name_con.insertByName(name, model)

            # reference the control by name
            ctrl_con = mLo.Lo.qi(XControlContainer, dialog_ctrl, True)

            # use the model's name to get its view inside the dialog
            result = cast("UnoControlTimeField", ctrl_con.getControl(name))
            cls._set_size_pos(result, x, y, width, height)
            ctl = CtlTimeField(result)
            ctl.time_format = time_format
            return ctl
        except Exception as e:
            raise mEx.DialogError(f"Could not create text field control: {e}") from e

    # endregion    add components to a dialog

    @staticmethod
    def _set_size_pos(ctl: XWindow, x: int = -1, y: int = -1, width: int = -1, height: int = -1) -> None:
        """
        Set Position and size for a control.

        |lo_safe|

        Args:
            ctl (XWindow): Control that implements XWindow
            x (int, optional): X Position. Defaults to -1.
            y (int, optional): Y Position. Defaults to -1.
            width (int, optional): Width. Defaults to -1.
            height (int, optional): Height. Defaults to -1.
        """
        if x < 0 and y < 0 and width < 0 and height < 0:
            return

        pos_size = None
        if x > -1 and y > -1 and width > -1 and height > -1:
            pos_size = PosSize.POSSIZE
        elif x > -1 and y > -1:
            pos_size = PosSize.POS
        elif width > -1 and height > -1:
            pos_size = PosSize.SIZE
        if pos_size is not None:
            ctl.setPosSize(x, y, width, height, pos_size)

    @classmethod
    def get_radio_group_value(cls, dialog_ctrl: XControl, radio_button: str) -> List[CtlRadioButton]:
        """
        Get a radio button group. Similar to :py:meth:`~.dialogs.Dialogs.find_radio_siblings` but also includes first radio button.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            radio_button (str): Name of the first radio button of the group.

        Returns:
            Any: Value of the selected radio button.

        See Also:
            :py:meth:`~.dialogs.Dialogs.find_radio_siblings`.
        """
        result: List[CtlRadioButton] = []
        ctl = cls.find_control(dialog_ctrl, radio_button)
        if not ctl:
            return result
        first_radio_btn = cast(CtlRadioButton, cls.get_dialog_control_instance(ctl))
        if first_radio_btn is None:
            return result
        result.append(first_radio_btn)
        result.extend(cls.find_radio_siblings(dialog_ctrl, radio_button))
        return result

    @classmethod
    def find_radio_siblings(cls, dialog_ctrl: XControl, radio_button: str) -> List[CtlRadioButton]:
        """
        Given the name of the first radio button of a group, return all the controls of the group.

        For dialogs, radio buttons are considered of the same group when their tab indexes are contiguous.

        |lo_unsafe|

        Args:
            dialog_ctrl (XControl): Control.
            radio_button (str): Specifies the exact name of the 1st radio button of the group.

        Returns:
            List[CtlRadioButton]: List of the names of the 1st and the next radio buttons.
            belonging to the same group in their tab index order. does not include the first button.

        See Also:
            :py:meth:`~.dialogs.Dialogs.get_radio_group_value`.
        """
        ctl = cls.find_control(dialog_ctrl, radio_button)
        result: List[CtlRadioButton] = []
        if not ctl:
            return result
        first_radio_btn = cast(CtlRadioButton, cls.get_dialog_control_instance(ctl))
        if first_radio_btn is None:
            return result
        if first_radio_btn.get_control_kind() != DialogControlKind.RADIO_BUTTON:
            return result

        tab_index = first_radio_btn.tab_index
        next_tab_index = tab_index + 1
        controls = cls.find_controls(dialog_ctrl, CtlRadioButton)
        if not controls:
            return result
        # controls.pop(0)
        for ctrl in controls:
            if ctrl.tab_index == next_tab_index:
                result.append(ctrl)
                next_tab_index += 1
        return result
