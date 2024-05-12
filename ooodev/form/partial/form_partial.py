"""Partial Class for Form Component. Not intended to be used directly."""

# region Imports
from __future__ import annotations
from typing import Any, cast, Iterable, TYPE_CHECKING
import datetime

import uno
from com.sun.star.form import XForms
from com.sun.star.container import XChild
from com.sun.star.awt import XControl
from com.sun.star.awt import XControlModel
from com.sun.star.drawing import XDrawPage

from ooo.dyn.text.text_content_anchor_type import TextContentAnchorType

from ooodev.exceptions import ex as mEx
from ooodev.form import forms as mForms
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.kind.border_kind import BorderKind
from ooodev.utils.kind.date_format_kind import DateFormatKind
from ooodev.utils.kind.orientation_kind import OrientationKind
from ooodev.utils.kind.state_kind import StateKind
from ooodev.utils.kind.time_format_kind import TimeFormatKind
from ooodev.utils.kind.tri_state_kind import TriStateKind
from ooodev.form.controls.form_ctl_button import FormCtlButton
from ooodev.form.controls.form_ctl_check_box import FormCtlCheckBox
from ooodev.form.controls.form_ctl_combo_box import FormCtlComboBox
from ooodev.form.controls.form_ctl_currency_field import FormCtlCurrencyField
from ooodev.form.controls.form_ctl_date_field import FormCtlDateField
from ooodev.form.controls.form_ctl_file import FormCtlFile
from ooodev.form.controls.form_ctl_formatted_field import FormCtlFormattedField
from ooodev.form.controls.form_ctl_grid import FormCtlGrid
from ooodev.form.controls.form_ctl_group_box import FormCtlGroupBox
from ooodev.form.controls.form_ctl_image_button import FormCtlImageButton
from ooodev.form.controls.form_ctl_fixed_text import FormCtlFixedText
from ooodev.form.controls.form_ctl_hidden import FormCtlHidden
from ooodev.form.controls.form_ctl_list_box import FormCtlListBox
from ooodev.form.controls.form_ctl_navigation_tool_bar import FormCtlNavigationToolBar
from ooodev.form.controls.form_ctl_numeric_field import FormCtlNumericField
from ooodev.form.controls.form_ctl_pattern_field import FormCtlPatternField
from ooodev.form.controls.form_ctl_radio_button import FormCtlRadioButton
from ooodev.form.controls.form_ctl_rich_text import FormCtlRichText
from ooodev.form.controls.form_ctl_scroll_bar import FormCtlScrollBar
from ooodev.form.controls.form_ctl_spin_button import FormCtlSpinButton
from ooodev.form.controls.form_ctl_submit_button import FormCtlSubmitButton
from ooodev.form.controls.form_ctl_text_field import FormCtlTextField
from ooodev.form.controls.form_ctl_time_field import FormCtlTimeField
from ooodev.form.controls.database.form_ctl_db_check_box import FormCtlDbCheckBox
from ooodev.form.controls.database.form_ctl_db_combo_box import FormCtlDbComboBox
from ooodev.form.controls.database.form_ctl_db_currency_field import FormCtlDbCurrencyField
from ooodev.form.controls.database.form_ctl_db_date_field import FormCtlDbDateField
from ooodev.form.controls.database.form_ctl_db_formatted_field import FormCtlDbFormattedField
from ooodev.form.controls.database.form_ctl_db_list_box import FormCtlDbListBox
from ooodev.form.controls.database.form_ctl_db_numeric_field import FormCtlDbNumericField
from ooodev.form.controls.database.form_ctl_db_pattern_field import FormCtlDbPatternField
from ooodev.form.controls.database.form_ctl_db_radio_button import FormCtlDbRadioButton
from ooodev.form.controls.database.form_ctl_db_text_field import FormCtlDbTextField
from ooodev.form.controls.database.form_ctl_db_time_field import FormCtlDbTimeField

if TYPE_CHECKING:
    from com.sun.star.form.component import Form
    from com.sun.star.lang import XComponent
    from com.sun.star.drawing import XShape
    from ooodev.form.controls.form_ctl_base import FormCtlBase
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.units.unit_obj import UnitT
    from ooodev.proto.style_obj import StyleT
    from ooodev.utils.type_var import PathOrStr


# endregion Imports


class FormPartial:
    """
    Method for adding controls and other form elements to a form.
    """

    def __init__(
        self, owner: ComponentT, draw_page: XDrawPage, component: Form, lo_inst: LoInst | None = None
    ) -> None:
        """
        Constructor

        Args:
            owner (ComponentT): Class that owns this component.
            draw_page (XDrawPage): Draw Page
            component (Form): Form component
        """
        if lo_inst is None:
            self.__lo_inst = mLo.Lo.current_lo
        else:
            self.__lo_inst = lo_inst
        assert mInfo.Info.support_service(
            component, "com.sun.star.form.component.Form"
        ), "component must support com.sun.star.form.component.Form service"
        self.__owner = owner
        self.__component = component
        self.__draw_page = draw_page
        with LoContext(self.__lo_inst):
            forms = mLo.Lo.qi(XForms, component.getParent(), True)
            self.__doc = cast("XComponent", mLo.Lo.qi(XChild, forms, True).getParent())

    # region Insert Control Methods

    def insert_control_button(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        label: str = "",
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlButton:
        """
        Inserts a button control.

        By Default the button has tab stop and does focus on click.

        Args:
            doc (XComponent): Component
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int, UnitT, optional): Button Width.
            height (int, UnitT, optional): Button Height. Defaults to ``6`` mm.
            label (str, optional): Button label (text).
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            parent_form (XNameContainer, optional): Parent form in which to add control.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlButton: Button Control
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            result = mForms.Forms.insert_control_button(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                label=label,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return result

    def insert_control_check_box(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        label: str = "",
        tri_state: bool = True,
        state: TriStateKind = TriStateKind.NOT_CHECKED,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlCheckBox:
        """
        Inserts a check box control into the form.

        Args:
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int | UnitT): Width
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            label (str, optional): Label (text) to assign to checkbox.
            tri_state (TriStateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``TriStateKind.CHECKED``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlCheckBox: Checkbox Control

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``TriStateKind`` can be imported from ``ooodev.utils.kind.tri_state_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            result = mForms.Forms.insert_control_check_box(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                label=label,
                tri_state=tri_state,
                state=state,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return result

    def insert_control_combo_box(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        entries: Iterable[str] | None = None,
        max_text_len: int = 0,
        drop_down: bool = True,
        read_only: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlComboBox:
        """
        Inserts a ComboBox control into the form.

        Args:
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int | UnitT): Width
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            entries (Iterable[str], optional): Combo box entries
            tri_state (TriStateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``TriStateKind.CHECKED``.
            max_text_len (int, optional): Specifies the maximum character count, There's no limitation, if set to 0. Defaults to ``0``.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``True``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlComboBox: ComboBox Control

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            result = mForms.Forms.insert_control_combo_box(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                entries=entries,
                max_text_len=max_text_len,
                drop_down=drop_down,
                read_only=read_only,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return result

    def insert_control_currency_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlCurrencyField:
        """
        Inserts a currency field control into the form.

        Args:
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int | UnitT): Width
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlCurrencyField: Currency Field Control

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            result = mForms.Forms.insert_control_currency_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                increment=increment,
                accuracy=accuracy,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return result

    def insert_control_date_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_date: datetime.datetime = datetime.datetime(1900, 1, 1, 0, 0, 0, 0),
        max_date: datetime.datetime = datetime.datetime(2200, 12, 31, 0, 0, 0, 0),
        drop_down: bool = True,
        date_format: DateFormatKind = DateFormatKind.SYSTEM_SHORT,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDateField:
        """
        Inserts a Date field control into the form.

        Args:
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int | UnitT): Width
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            date_value (datetime.datetime | None, optional): Specifics control datetime. Defaults to ``None``.
            min_date (datetime.datetime, optional): Specifics control min datetime. Defaults to ``datetime(1900, 1, 1, 0, 0, 0, 0)``.
            max_date (datetime.datetime, optional): Specifics control Min datetime. Defaults to ``datetime(2200, 12, 31, 0, 0, 0, 0)``.
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to ``True``.
            date_format (DateFormatKind, optional): Date format. Defaults to ``DateFormatKind.SYSTEM_SHORT``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDateField: Date Field Control

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``DateFormatKind`` can be imported from ``ooodev.utils.kind.date_format_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_date_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_date=min_date,
                max_date=max_date,
                drop_down=drop_down,
                date_format=date_format,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_file(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlFile:
        """
        Inserts a file control.

        Args:
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int, UnitT, optional): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlFile: File Control
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_file(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_formatted_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlFormattedField:
        """
        Inserts a currency field control into the form.

        Args:
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int | UnitT): Width
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlFormattedField: Currency Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_formatted_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_group_box(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        label: str = "",
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlGroupBox:
        """
        Inserts a Groupbox control into the form.

        Args:
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int | UnitT): Width
            height (int, UnitT): Height.
            label (str, optional): Groupbox label.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlGroupBox: Groupbox Control
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_group_box(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                label=label,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_grid(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        label: str = "",
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlGrid:
        """
        Inserts a Grid control.

        Args:
            x (int | UnitT): X Coordinate
            y (int | UnitT): Y Coordinate
            width (int | UnitT): Width
            height (int, UnitT): Height.
            label (str, optional): Grid label.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlGrid: Grid Control
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_grid(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                label=label,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_hidden(self, *, name: str = "", **kwargs: Any) -> FormCtlHidden:
        """
        Inserts a Hidden control into the form.

        Args:
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.

        Returns:
            FormCtlHidden: Hidden Control.
        """
        # **kwargs are just for backwards compatibility.
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_hidden(
                name=name,
                parent_form=self.__component,
            )
        return results

    def insert_control_image_button(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        image_url: PathOrStr = "",
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlImageButton:
        """
        Inserts an Image Button control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT): Height.
            image_url (PathOrStr, optional): Image URL. When setting the value it can be a string or a Path object.
                If a string is passed it can be a URL or a path to a file.
                Value such as ``file:///path/to/image.png`` and ``/path/to/image.png`` are valid.
                Relative paths are supported.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlImageButton: Image Button Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_image_button(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                image_url=image_url,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_label(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        label: str,
        height: int | UnitT = 6,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlFixedText:
        """
        Inserts a Label control.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int, UnitT, optional): Width.
            label (str): Contents of label.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlFixedText: Label Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_label(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                label=label,
                height=height,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_list_box(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        entries: Iterable[str] | None = None,
        drop_down: bool = True,
        read_only: bool = False,
        line_count: int = 5,
        multi_select: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlListBox:
        """
        Inserts a ListBox control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT): Height.
            entries (Iterable[str], optional): Combo box entries
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``True``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            line_count (int, optional): Specifies the number of lines to display. Defaults to ``5``.
            multi_select (int, optional): Specifies if multiple entries can be selected. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlListBox: ListBox Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_list_box(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                entries=entries,
                drop_down=drop_down,
                read_only=read_only,
                line_count=line_count,
                multi_select=multi_select,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_navigation_toolbar(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlNavigationToolBar:
        """
        Inserts a Navigation Toolbar control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT): Height.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlNavigationToolBar: Navigation Toolbar Control
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_navigation_toolbar(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_numeric_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlNumericField:
        """
        Inserts a Numeric field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlNumericField: Numeric Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_numeric_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                increment=increment,
                accuracy=accuracy,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_pattern_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        edit_mask: str = "",
        literal_mask: str = "",
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlPatternField:
        """
        Inserts a Pattern field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            edit_mask (str, optional): Specifies a character code that determines what the user may enter. Defaults to ``""``.
            literal_mask (str, optional): Specifies the initial values that are displayed in the pattern field. Defaults to ``""``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlPatternField: Pattern Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_pattern_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                edit_mask=edit_mask,
                literal_mask=literal_mask,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_radio_button(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        label: str = "",
        state: StateKind = StateKind.NOT_CHECKED,
        multiline: bool = False,
        border: BorderKind = BorderKind.NONE,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlRadioButton:
        """
        Inserts a radio button control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            label (str, optional): Label (text) of control.
            anchor_type (TextContentAnchorType | None, optional): _description_. Defaults to None.
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
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_radio_button(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                label=label,
                state=state,
                multiline=multiline,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_rich_text(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlRichText:
        """
        Inserts a Rich Text control.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int, UnitT, optional): Width.
            height (int, UnitT, optional): Height.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlRichText: Rich Text Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_rich_text(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_scroll_bar(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_value: int = 0,
        max_value: int = 100,
        orientation: OrientationKind = OrientationKind.HORIZONTAL,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlScrollBar:
        """
        Inserts a Scrollbar control.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int, UnitT, optional): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``100``.
            orientation (OrientationKind, optional): Orientation. Defaults to ``OrientationKind.HORIZONTAL``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlScrollBar: Scrollbar Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``OrientationKind`` can be imported from ``ooodev.utils.kind.orientation_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_scroll_bar(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                orientation=orientation,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_spin_button(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        value: int = 0,
        min_value: int = -1000000,
        max_value: int = 1000000,
        increment: int = 1,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlSpinButton:
        """
        Inserts a Spin Button control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            value (int, optional): Specifies the initial value of the control. Defaults to ``0``.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlSpinButton: Spin Button Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_spin_button(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                value=value,
                height=height,
                min_value=min_value,
                max_value=max_value,
                increment=increment,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_submit_button(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlSubmitButton:
        """
        Inserts a submit button control.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int, UnitT, optional): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlSubmitButton: Submit Button Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_submit_button(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_text_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        text: str = "",
        echo_char: str = "",
        border: BorderKind = BorderKind.NONE,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlTextField:
        """
        Inserts a Text field control.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int, UnitT, optional): Width.
            height (int, UnitT, optional): Height.
            text (str, optional): Text value.
            echo_char (str, optional): Character used for masking. Must be a single character.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlTextField: Text Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_text_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                text=text,
                height=height,
                echo_char=echo_char,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_control_time_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        time_value: datetime.time | None = None,
        min_time: datetime.time = datetime.time(0, 0, 0, 0),
        max_time: datetime.time = datetime.time(23, 59, 59, 999_999),
        time_format: TimeFormatKind = TimeFormatKind.SHORT_24H,
        spin_button: bool = True,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlTimeField:
        """
        Inserts a Time field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            time_value (datetime.time | None, optional): Specifics the control time. Defaults to ``None``.
            min_time (datetime.time, optional): Specifics control min time. Defaults to ``time(0, 0, 0, 0)``.
            max_time (datetime.time, optional): Specifics control min time. Defaults to a ``time(23, 59, 59, 999_999)``.
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to ``True``.
            time_format (TimeFormatKind, optional): Date format. Defaults to ``TimeFormatKind.SHORT_24H``.
            pin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``True``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlTimeField: Time Field Control.

        Hint:
            - ``BorderKind`` can be imported from ``ooodev.utils.kind.border_kind``.
            - ``TimeFormatKind`` can be imported from ``ooodev.utils.kind.time_format_kind``.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_control_time_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
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
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_check_box(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        tri_state: bool = True,
        state: TriStateKind = TriStateKind.CHECKED,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbCheckBox:
        """
        Inserts a database check box control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            anchor_type (TextContentAnchorType | None, optional): _description_. Defaults to None.
            tri_state (TriStateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``TriStateKind.CHECKED``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbCheckBox: Database Checkbox Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_check_box(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                tri_state=tri_state,
                state=state,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_combo_box(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        entries: Iterable[str] | None = None,
        max_text_len: int = 0,
        drop_down: bool = True,
        read_only: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbComboBox:
        """
        Inserts a  Database ComboBox control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            entries (Iterable[str], optional): Combo box entries
            tri_state (TriStateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``TriStateKind.CHECKED``.
            max_text_len (int, optional): Specifies the maximum character count, There's no limitation, if set to 0. Defaults to ``0``.
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``True``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbComboBox: Database ComboBox Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_combo_box(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                entries=entries,
                height=height,
                max_text_len=max_text_len,
                drop_down=drop_down,
                read_only=read_only,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_currency_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbCurrencyField:
        """
        Inserts a database currency field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbCurrencyField: Database Currency Field Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_currency_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                increment=increment,
                accuracy=accuracy,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_date_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_date: datetime.datetime = datetime.datetime(1900, 1, 1, 0, 0, 0, 0),
        max_date: datetime.datetime = datetime.datetime(2200, 12, 31, 0, 0, 0, 0),
        drop_down: bool = True,
        date_format: DateFormatKind = DateFormatKind.SYSTEM_SHORT,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbDateField:
        """
        Inserts a Database Date field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            date_value (datetime.datetime | None, optional): Specifics control datetime. Defaults to ``None``.
            min_date (datetime.datetime, optional): Specifics control min datetime. Defaults to ``datetime(1900, 1, 1, 0, 0, 0, 0)``.
            max_date (datetime.datetime, optional): Specifics control Min datetime. Defaults to ``datetime(2200, 12, 31, 0, 0, 0, 0)``.
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to ``True``.
            date_format (DateFormatKind, optional): Date format. Defaults to ``DateFormatKind.SYSTEM_SHORT``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbDateField: Database Date Field Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_date_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_date=min_date,
                max_date=max_date,
                drop_down=drop_down,
                date_format=date_format,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_formatted_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbFormattedField:
        """
        Inserts a Database currency field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbFormattedField: Database Currency Field Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_formatted_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_list_box(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        entries: Iterable[str] | None = None,
        drop_down: bool = True,
        read_only: bool = False,
        line_count: int = 5,
        multi_select: bool = False,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbListBox:
        """
        Inserts a Database ListBox control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT): Height.
            entries (Iterable[str], optional): Combo box entries
            drop_down (bool, optional): Specifies if the control has a drop down button. Defaults to ``True``.
            read_only (bool, optional): Specifies that the content of the control cannot be modified by the user. Defaults to ``False``.
            line_count (int, optional): Specifies the number of lines to display. Defaults to ``5``.
            multi_select (int, optional): Specifies if multiple entries can be selected. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbListBox: Database ListBox Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_list_box(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                entries=entries,
                height=height,
                drop_down=drop_down,
                read_only=read_only,
                line_count=line_count,
                multi_select=multi_select,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_numeric_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        min_value: float = -1000000.0,
        max_value: float = 1000000.0,
        spin_button: bool = False,
        increment: int = 1,
        accuracy: int = 2,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbNumericField:
        """
        Inserts a Database Numeric field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            min_value (float, optional): Specifies the smallest value that can be entered in the control. Defaults to ``-1000000.0``.
            max_value (float, optional): Specifies the largest value that can be entered in the control. Defaults to ``1000000.0``.
            spin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``False``.
            increment (int, optional): The step when the spin button is pressed. Defaults to ``1``.
            accuracy (int, optional): Specifies the decimal accuracy. Default is ``2`` decimal digits
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbNumericField: Database Numeric Field Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_numeric_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                min_value=min_value,
                max_value=max_value,
                spin_button=spin_button,
                increment=increment,
                accuracy=accuracy,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_pattern_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        edit_mask: str = "",
        literal_mask: str = "",
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbPatternField:
        """
        Inserts a Database Pattern field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            edit_mask (str, optional): Specifies a character code that determines what the user may enter. Defaults to ``""``.
            literal_mask (str, optional): Specifies the initial values that are displayed in the pattern field. Defaults to ``""``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbPatternField: Database Pattern Field Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_pattern_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                edit_mask=edit_mask,
                literal_mask=literal_mask,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_radio_button(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        state: StateKind = StateKind.NOT_CHECKED,
        multiline: bool = False,
        border: BorderKind = BorderKind.NONE,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbRadioButton:
        """
        Inserts a Database radio button control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            anchor_type (TextContentAnchorType | None, optional): _description_. Defaults to None.
            tri_state (StateKind, optional): Specifies that the control may have the state "don't know". Defaults to ``True``.
            state (TriStateKind, optional): Specifies the state of the control.Defaults to ``StateKind.NOT_CHECKED``.
            multiline (bool, optional): Specifies if the control can display multiple lines of text. Defaults to ``False``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbRadioButton: Database Radio Button Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_radio_button(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                height=height,
                state=state,
                multiline=multiline,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_text_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT,
        text: str = "",
        echo_char: str = "",
        border: BorderKind = BorderKind.NONE,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbTextField:
        """
        Inserts a Database Text field control.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int, UnitT, optional): Width.
            height (int, UnitT, optional): Height.
            text (str, optional): Text value.
            echo_char (str, optional): Character used for masking. Must be a single character.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.NONE``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlDbTextField: Database Text Field Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_text_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
                x=x,
                y=y,
                width=width,
                text=text,
                height=height,
                echo_char=echo_char,
                border=border,
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    def insert_db_control_time_field(
        self,
        *,
        x: int | UnitT,
        y: int | UnitT,
        width: int | UnitT,
        height: int | UnitT = 6,
        time_value: datetime.time | None = None,
        min_time: datetime.time = datetime.time(0, 0, 0, 0),
        max_time: datetime.time = datetime.time(23, 59, 59, 999_999),
        time_format: TimeFormatKind = TimeFormatKind.SHORT_24H,
        spin_button: bool = True,
        border: BorderKind = BorderKind.BORDER_3D,
        anchor_type: TextContentAnchorType = TextContentAnchorType.AT_PARAGRAPH,
        name: str = "",
        styles: Iterable[StyleT] | None = None,
    ) -> FormCtlDbTimeField:
        """
        Inserts a Database Time field control into the form.

        Args:
            x (int | UnitT): X Coordinate.
            y (int | UnitT): Y Coordinate.
            width (int | UnitT): Width.
            height (int, UnitT, optional): Height. Defaults to ``6`` mm.
            time_value (datetime.time | None, optional): Specifics the control time. Defaults to ``None``.
            min_time (datetime.time, optional): Specifics control min time. Defaults to ``time(0, 0, 0, 0)``.
            max_time (datetime.time, optional): Specifics control min time. Defaults to a ``time(23, 59, 59, 999_999)``.
            drop_down (bool, optional): Specifies if the control is a dropdown. Defaults to ``True``.
            time_format (TimeFormatKind, optional): Date format. Defaults to ``TimeFormatKind.SHORT_24H``.
            pin_button (bool, optional): When ``True``, a spin button is present. Defaults to ``True``.
            border (BorderKind, optional): Border option. Defaults to ``BorderKind.BORDER_3D``.
            anchor_type (TextContentAnchorType, optional): Control Anchor Type. Defaults to ``TextContentAnchorType.AT_PARAGRAPH``
            name (str, optional): Name of control. Must be a unique name. If empty, a unique name is generated.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlTimeField: Database Time Field Control.
        """
        # when control is created, it will automatically get the same lo instance as LoContext.
        with LoContext(self.__lo_inst):
            results = mForms.Forms.insert_db_control_time_field(
                doc=self.__doc,
                draw_page=self.__draw_page,
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
                anchor_type=anchor_type,
                name=name,
                parent_form=self.__component,
                styles=styles,
            )
        return results

    # endregion Insert Control Methods

    # region Other Form Methods
    def bind_form_to_sql(self, src_name: str, cmd: str) -> None:
        """
        Bind the form to the database in the ``src_name`` URL, and send a SQL cmd

        Args:
            src_name (str): Source Name URL
            cmd (str): Command

        Returns:
            None:
        """
        with LoContext(self.__lo_inst):
            mForms.Forms.bind_form_to_sql(self.__component, src_name, cmd)

    def bind_form_to_table(self, src_name: str, tbl_name: str) -> None:
        """
        Bind the form to the database in the src_name URL

        Args:
            src_name (str): Source Name URL
            tbl_name (str): Table Name

        Returns:
            None:
        """
        with LoContext(self.__lo_inst):
            mForms.Forms.bind_form_to_table(self.__component, src_name, tbl_name)

    def get_control(self, ctl_model: XControlModel) -> XControl:
        """
        Gets the control from the specified control model.

        Args:
            ctl_model (XControlModel): Control Model

        Raises:
            Exception: If unable to get control

        Returns:
            XControl: Control
        """
        with LoContext(self.__lo_inst):
            results = mForms.Forms.get_control(self.__doc, ctl_model)
        return results

    def get_control_model(self, ctl_name: str) -> XControlModel:
        """
        Gets Control Model by Name

        Args:
            ctl_name (str): Name of control

        Raises:
            MissingNameError: If control not found

        Returns:
            XControlModel | None: Control Model if found; Otherwise, None
        """
        if self.__component.hasByName(ctl_name):
            with LoContext(self.__lo_inst):
                result = mForms.Forms.get_control_model(self.__doc, ctl_name)
            if result is None:
                raise mEx.MissingNameError(f"Control '{ctl_name}' not found")
            return result
        else:
            raise mEx.MissingNameError(f"Control '{ctl_name}' not found")

    def get_control_index(self, ctl: FormCtlBase | XControlModel) -> int:
        """
        Get the index of the control within the form.

        Args:
            ctl (FormCtlBase, XControlModel): Control object.

        Returns:
            int: Control Index within the form or ``-1`` if not found.

        .. versionadded:: 0.38.0
        """
        return mForms.Forms.get_control_index(self.__component, ctl)

    def find_shape_for_control(self, ctl: FormCtlBase | XControlModel) -> XShape | None:
        """
        Find the shape for a control.

        Args:
            control (FormCtlBase | XControlModel): control to find shape for.

        Returns:
            XShape | None: Shape for the control or ``None`` if not found.

        .. versionadded:: 0.38.0
        """
        return mForms.Forms.find_shape_for_control(self.__draw_page, ctl)

    # endregion Other Form Methods
