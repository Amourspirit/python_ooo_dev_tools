# region Imports
from __future__ import annotations
from abc import abstractmethod
from typing import cast, Any, Iterable, TYPE_CHECKING, Tuple
import datetime
import uno
from com.sun.star.drawing import XControlShape
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.kind.date_format_kind import DateFormatKind
from ooodev.utils.kind.tri_state_kind import TriStateKind
from ooodev.utils.kind.state_kind import StateKind
from ooodev.utils.kind.orientation_kind import OrientationKind
from ooodev.utils.kind.time_format_kind import TimeFormatKind
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.loader import lo as mLo
from ooodev.form import forms as mForms
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.events.partial.events_partial import EventsPartial


if TYPE_CHECKING:
    from com.sun.star.sheet import Shape  # service
    from ooodev.calc.calc_form import CalcForm
    from ooodev.form.controls.form_ctl_button import FormCtlButton
    from ooodev.form.controls.form_ctl_check_box import FormCtlCheckBox
    from ooodev.form.controls.form_ctl_combo_box import FormCtlComboBox
    from ooodev.form.controls.form_ctl_currency_field import FormCtlCurrencyField
    from ooodev.form.controls.form_ctl_date_field import FormCtlDateField
    from ooodev.form.controls.form_ctl_file import FormCtlFile
    from ooodev.form.controls.form_ctl_fixed_text import FormCtlFixedText
    from ooodev.form.controls.form_ctl_formatted_field import FormCtlFormattedField
    from ooodev.form.controls.form_ctl_text_field import FormCtlTextField
    from ooodev.form.controls.form_ctl_grid import FormCtlGrid
    from ooodev.form.controls.form_ctl_list_box import FormCtlListBox
    from ooodev.form.controls.form_ctl_group_box import FormCtlGroupBox
    from ooodev.form.controls.form_ctl_image_button import FormCtlImageButton
    from ooodev.form.controls.form_ctl_list_box import FormCtlListBox
    from ooodev.form.controls.form_ctl_numeric_field import FormCtlNumericField
    from ooodev.form.controls.form_ctl_pattern_field import FormCtlPatternField
    from ooodev.form.controls.form_ctl_radio_button import FormCtlRadioButton
    from ooodev.form.controls.form_ctl_scroll_bar import FormCtlScrollBar
    from ooodev.form.controls.form_ctl_spin_button import FormCtlSpinButton
    from ooodev.form.controls.form_ctl_time_field import FormCtlTimeField
    from ooodev.form.controls.form_ctl_rich_text import FormCtlRichText

    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.type_var import PathOrStr

# endregion Imports

# pylint: disable=super-init-not-called


class SheetControlBase(LoInstPropsPartial, CalcSheetPropPartial, EventsPartial, TheDictionaryPartial):
    """A partial class for a cell control."""

    def __init__(self, calc_obj: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._calc_obj = calc_obj
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        EventsPartial.__init__(self)
        TheDictionaryPartial.__init__(self)
        self._init_calc_sheet_prop()
        self._current_control = NULL_OBJ

    # region Abstract Methods
    @abstractmethod
    def _init_calc_sheet_prop(self) -> None:
        raise NotImplementedError

    @abstractmethod
    def _get_pos_size(self) -> Tuple[int, int, int, int]:
        raise NotImplementedError

    # endregion Abstract Methods

    # region protected methods

    def _set_shape_props(self, shape: Shape) -> None:
        event_data = {"Anchor": self.calc_obj.component, "ResizeWithCell": True, "MoveProtect": True}
        eargs = CancelEventArgs(source=shape)
        eargs.event_data = event_data
        self.on_setting_shape_props(eargs)
        if eargs.cancel:
            return
        for key, value in eargs.event_data.items():
            setattr(shape, key, value)

    def _get_form(self) -> CalcForm:
        sheet = self.calc_sheet
        if len(sheet.draw_page.forms) == 0:
            sheet.draw_page.forms.add_form("Form1")
        return sheet.draw_page.forms[0]

    def _return_control(self, control):
        shape = cast("Shape", control.control_shape)
        self._set_shape_props(shape)
        self._current_control = control
        eargs = EventArgs(source=self)
        eargs.event_data = {"control": control}
        self.on_insert_control(eargs)
        return control

    def _find_current_control(self) -> Any:
        # pylint: disable=import-outside-toplevel
        ps = self._get_pos_size()
        cargs = CancelEventArgs(source=self)
        cargs.event_data = {"pos_size": ps, "find_by_pos": True}
        self.on_finding_control(cargs)
        if cargs.cancel:
            return None

        sheet = self.calc_sheet
        if len(sheet.draw_page.forms) == 0:
            return None

        if cargs.event_data.get("find_by_pos", True):
            shape = sheet.draw_page.find_shape_at_position(ps[0], ps[1])
        else:
            shape = sheet.draw_page.find_shape_at_position_size(ps[0], ps[1], ps[2], ps[3])
        if shape is None:
            return None
        x_shape = mLo.Lo.qi(XControlShape, shape.component)
        if x_shape is None:
            return None

        ctl = x_shape.getControl()
        if ctl is None:
            return None
        from ooodev.form.controls.from_control_factory import FormControlFactory

        factory = FormControlFactory(draw_page=sheet.draw_page.component, lo_inst=self.lo_inst)
        return factory.get_control_from_model(ctl)

    def _remove_control(self, ctl: Any) -> None:
        if ctl is None:
            return
        sheet = self.calc_sheet
        dp = sheet.draw_page
        shape = ctl.control_shape
        dp.remove(shape)

    # endregion protected methods

    # region Event Handlers

    def on_insert_control(self, event_args: EventArgs) -> None:
        """
        Event handler for insert control.

        Triggers the ``insert_control`` event.
        """
        self.trigger_event("insert_control", event_args)

    def on_finding_control(self, event_args: CancelEventArgs) -> None:
        """
        Event handler for finding control.

        Triggers the ``finding_control`` event.
        """
        self.trigger_event("finding_control", event_args)

    def on_setting_shape_props(self, event_args: CancelEventArgs) -> None:
        """
        Event handler for setting shape properties.

        Triggers the ``setting_shape_props`` event.
        """
        self.trigger_event("setting_shape_props", event_args)

    # endregion Event Handlers

    # region Insert Controls

    def insert_control_button(
        self,
        label: str = "",
        *,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlButton:
        """
        Inserts a button control into sheet.

        By Default the button has tab stop and does focus on click.

        Args:
            label (str, optional): Button label (text).
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlButton: Button Control
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_button(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                label=label,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_check_box(
        self,
        label: str = "",
        *,
        name: str = "",
        tri_state: bool = True,
        state: TriStateKind = TriStateKind.NOT_CHECKED,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlCheckBox:
        """
        Inserts a check box control into the sheet.

        Args:
            label (str, optional): Label (text) to assign to checkbox.
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            tri_state (TriStateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``TriStateKind.CHECKED``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlCheckBox: Checkbox Control

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``TriStateKind`` can be imported from ``ooodev.utils.kind.tri_state_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_check_box(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                label=label,
                tri_state=tri_state,
                state=state,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_combo_box(
        self,
        *,
        name: str = "",
        entries: Iterable[str] | None = None,
        max_text_len: int = 0,
        drop_down: bool = True,
        read_only: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlComboBox:
        """
        Inserts a ComboBox control into the sheet.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            entries (Iterable[str], optional): Combo box entries
            tri_state (TriStateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``TriStateKind.CHECKED``.
            max_text_len (int, optional): Specifies the maximum character count, There's no limitation, if set to 0. Defaults to ``0``.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``True``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlComboBox: ComboBox Control

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_combo_box(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                entries=entries,
                max_text_len=max_text_len,
                drop_down=drop_down,
                read_only=read_only,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_currency_field(
        self,
        *,
        name: str = "",
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlCurrencyField:
        """
        Inserts a currency field control into the sheet.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlCurrencyField: Currency Field Control

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_currency_field(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                increment=increment,
                accuracy=accuracy,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_date_field(
        self,
        *,
        name: str = "",
        min_date: datetime.datetime = datetime.datetime(1900, 1, 1, 0, 0, 0, 0),
        max_date: datetime.datetime = datetime.datetime(2200, 12, 31, 0, 0, 0, 0),
        drop_down: bool = True,
        date_format: DateFormatKind = DateFormatKind.SYSTEM_SHORT,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDateField:
        """
        Inserts a Date field control into the form.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            date_value (datetime.datetime | None, optional): Specifics control datetime. Defaults to ``None``.
            min_date (datetime.datetime, optional): Specifics control min datetime. Defaults to ``datetime(1900, 1, 1, 0, 0, 0, 0)``.
            max_date (datetime.datetime, optional): Specifics control Min datetime. Defaults to ``datetime(2200, 12, 31, 0, 0, 0, 0)``.
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to ``True``.
            date_format (DateFormatKind, optional): Date format. Defaults to ``DateFormatKind.SYSTEM_SHORT``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDateField: Date Field Control

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``DateFormatKind`` can be imported from ``ooodev.utils.kind.date_format_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_date_field(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                min_date=min_date,
                max_date=max_date,
                drop_down=drop_down,
                date_format=date_format,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_file(
        self,
        *,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlFile:
        """
        Inserts a file control.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlFile: File Control.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_file(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_formatted_field(
        self,
        *,
        name: str = "",
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlFormattedField:
        """
        Inserts a currency field control into the form.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlFormattedField: Currency Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_formatted_field(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_group_box(
        self,
        label: str = "",
        *,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlGroupBox:
        """
        Inserts a Groupbox control into the form.

        Args:
            label (str, optional): Groupbox label.
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlGroupBox: Groupbox Control
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_group_box(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                label=label,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_grid(
        self,
        label: str = "",
        *,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlGrid:
        """
        Inserts a Grid control.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            label (str, optional): Grid label.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlGrid: Grid Control
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_grid(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                label=label,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_image_button(
        self,
        image_url: PathOrStr = "",
        *,
        name: str = "",
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlImageButton:
        """
        Inserts an Image Button control into the form.

        Args:
            image_url (PathOrStr, optional): Image URL. When setting the value it can be a string or a Path object.
                If a string is passed it can be a URL or a path to a file.
                Value such as ``file:///path/to/image.png`` and ``/path/to/image.png`` are valid.
                Relative paths are supported.
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlImageButton: Image Button Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_image_button(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                image_url=image_url,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_label(
        self,
        label: str,
        *,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlFixedText:
        """
        Inserts a Label control.

        Args:
            label (str): Contents of label.
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlFixedText: Label Control.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_label(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                label=label,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_list_box(
        self,
        *,
        entries: Iterable[str] | None = None,
        name: str = "",
        drop_down: bool = True,
        read_only: bool = False,
        line_count: int = 5,
        multi_select: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlListBox:
        """
        Inserts a ListBox control into the form.

        Args:
            entries (Iterable[str], optional): Combo box entries
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``True``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            line_count (int, optional): Specifies the number of lines to display. Defaults to ``5``.
            multi_select (int, optional): Specifies if multiple entries can be selected. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlListBox: ListBox Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_list_box(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                entries=entries,
                drop_down=drop_down,
                read_only=read_only,
                line_count=line_count,
                multi_select=multi_select,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_numeric_field(
        self,
        *,
        name: str = "",
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlNumericField:
        """
        Inserts a Numeric field control into the form.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlNumericField: Numeric Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_numeric_field(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                increment=increment,
                accuracy=accuracy,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_pattern_field(
        self,
        *,
        name: str = "",
        edit_mask: str = "",
        literal_mask: str = "",
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlPatternField:
        """
        Inserts a Pattern field control into the form.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            edit_mask (str, optional): Specifies a character code that determines what the user may enter. Defaults to ``""``.
            literal_mask (str, optional): Specifies the initial values that are displayed in the pattern field. Defaults to ``""``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlPatternField: Pattern Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_pattern_field(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                edit_mask=edit_mask,
                literal_mask=literal_mask,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_radio_button(
        self,
        label: str = "",
        *,
        state: StateKind = StateKind.NOT_CHECKED,
        multiline: bool = False,
        border: BorderKind = BorderKind.NONE,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlRadioButton:
        """
        Inserts a radio button control into the form.

        Args:
            label (str, optional): Label (text) of control.
            tri_state (StateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (StateKind, optional): Specifies the state of the control.Defaults to ``StateKind.NOT_CHECKED``.
            multiline (bool, optional): Specifies if the control can display multiple lines of text. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlRadioButton: Radio Button Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``StateKind`` can be imported from ``ooodev.utils.kind.state_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_radio_button(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                label=label,
                state=state,
                multiline=multiline,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_rich_text(
        self,
        *,
        name: str = "",
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlRichText:
        """
        Inserts a Rich Text control.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlRichText: Rich Text Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_rich_text(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_scroll_bar(
        self,
        *,
        name: str = "",
        min_value: int = 0,
        max_value: int = 100,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlScrollBar:
        """
        Inserts a Scrollbar control.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlScrollBar: Scrollbar Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_scroll_bar(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                min_value=min_value,
                max_value=max_value,
                orientation=orientation,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_spin_button(
        self,
        value: int = 0,
        *,
        name: str = "",
        min_value: int = -1000000,
        max_value: int = 1000000,
        increment: int = 1,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlSpinButton:
        """
        Inserts a Spin Button control into the form.

        Args:
            value (int, optional): Specifies the initial value of the control. Defaults to ``0``.
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlSpinButton: Spin Button Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_spin_button(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                value=value,
                min_value=min_value,
                max_value=max_value,
                increment=increment,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_text_field(
        self,
        text: str = "",
        *,
        name: str = "",
        echo_char: str = "",
        border: BorderKind = BorderKind.NONE,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlTextField:
        """
        Inserts a Text field control.

        Args:
            text (str, optional): Text value.
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            echo_char (str, optional): Character used for masking. Must be a single character.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlTextField: Text Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_text_field(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                text=text,
                echo_char=echo_char,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    def insert_control_time_field(
        self,
        time_value: datetime.time | None = None,
        *,
        name: str = "",
        min_time: datetime.time = datetime.time(0, 0, 0, 0),
        max_time: datetime.time = datetime.time(23, 59, 59, 999_999),
        time_format: TimeFormatKind = TimeFormatKind.SHORT_24H,
        spin_button: bool = True,
        border: BorderKind = BorderKind.BORDER_3D,
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlTimeField:
        """
        Inserts a Time field control into the form.

        Args:
            time_value (datetime.time | None, optional): Specifics the control time. Defaults to ``None``.
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            min_time (datetime.time, optional): Specifics control min time. Defaults to ``time(0, 0, 0, 0)``.
            max_time (datetime.time, optional): Specifics control min time. Defaults to a ``time(23, 59, 59, 999_999)``.
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to ``True``.
            time_format (TimeFormatKind, optional): Date format. Defaults to ``TimeFormatKind.SHORT_24H``.
            pin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``True``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlTimeField: Time Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``TimeFormatKind`` can be imported from ``ooodev.utils.kind.time_format_kind``.
        """
        x, y, width, height = self._get_pos_size()
        sheet = self.calc_sheet
        form = self._get_form()
        with LoContext(self.lo_inst):
            control = mForms.Forms.insert_control_time_field(
                doc=sheet.calc_doc.component,
                draw_page=sheet.draw_page.component,
                x=UnitMM100(x),
                y=UnitMM100(y),
                width=UnitMM100(width),
                height=UnitMM100(height),
                time_value=time_value,
                min_time=min_time,
                max_time=max_time,
                time_format=time_format,
                spin_button=spin_button,
                border=border,
                name=name,
                parent_form=form.component,
                styles=styles,
            )
        return self._return_control(control)

    # endregion Insert Controls

    # region Properties

    @property
    def calc_obj(self) -> Any:
        return self._calc_obj

    @property
    def current_control(self) -> Any:
        """
        Gets/Sets the control.

        When setting the control any previous control for the cell is removed.
        """
        # pylint: disable=import-outside-toplevel
        if self._current_control is NULL_OBJ:
            self._current_control = self._find_current_control()

        return self._current_control

    @current_control.setter
    def current_control(self, value: Any) -> None:
        """Sets the control."""
        # any old control should be removed before setting a new one
        if self._current_control is not NULL_OBJ and self._current_control is not None:
            self._remove_control(self._current_control)
        else:
            cc = self._find_current_control()
            if cc is not None:
                self._remove_control(cc)
        self._current_control = value

    # endregion Properties
