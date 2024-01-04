from __future__ import annotations
from typing import Any, cast, Iterable, TYPE_CHECKING
import uno
from com.sun.star.form import XForms
from com.sun.star.container import XChild

from ooodev.form import forms as mForms
from ooodev.utils import lo as mLo
from ooodev.utils import info as mInfo

# region Other Imports
import contextlib
import datetime
import uno

from com.sun.star.awt import XControl
from com.sun.star.awt import XControlModel
from com.sun.star.awt import XView
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XChild
from com.sun.star.container import XIndexContainer
from com.sun.star.container import XNameAccess
from com.sun.star.container import XNameContainer
from com.sun.star.container import XNamed
from com.sun.star.drawing import XControlShape
from com.sun.star.drawing import XDrawPage
from com.sun.star.drawing import XDrawPagesSupplier
from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.drawing import XShapes
from com.sun.star.form import XForm
from com.sun.star.form import XFormsSupplier
from com.sun.star.form import XGridColumnFactory
from com.sun.star.lang import XServiceInfo
from com.sun.star.script import XEventAttacherManager

from ooo.dyn.awt.point import Point
from ooo.dyn.awt.size import Size as UnoSize
from ooo.dyn.form.form_component_type import FormComponentType
from ooo.dyn.form.list_source_type import ListSourceType
from ooo.dyn.script.script_event_descriptor import ScriptEventDescriptor
from ooo.dyn.sdb.command_type import CommandType
from ooo.dyn.text.text_content_anchor_type import TextContentAnchorType


from ooodev.proto.style_obj import StyleT
from ooodev.utils import gen_util as gUtil
from ooodev.utils import gui as mGui
from ooodev.utils import info as mInfo
from ooodev.utils import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.date_format_kind import DateFormatKind as DateFormatKind
from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.utils.kind.language_kind import LanguageKind as LanguageKind
from ooodev.utils.kind.orientation_kind import OrientationKind as OrientationKind
from ooodev.utils.kind.state_kind import StateKind as StateKind
from ooodev.utils.kind.time_format_kind import TimeFormatKind as TimeFormatKind
from ooodev.utils.kind.tri_state_kind import TriStateKind as TriStateKind
from ..controls import FormCtlButton
from ..controls import FormCtlCheckBox
from ..controls import FormCtlComboBox
from ..controls import FormCtlCurrencyField
from ..controls import FormCtlDateField
from ..controls import FormCtlFile
from ..controls import FormCtlFormattedField
from ..controls import FormCtlGrid
from ..controls import FormCtlGroupBox
from ..controls import FormCtlImageButton
from ..controls import FormCtlFixedText
from ..controls import FormCtlHidden
from ..controls import FormCtlListBox
from ..controls import FormCtlNavigationToolBar
from ..controls import FormCtlNumericField
from ..controls import FormCtlPatternField
from ..controls import FormCtlRadioButton
from ..controls import FormCtlRichText
from ..controls import FormCtlScrollBar
from ..controls import FormCtlSpinButton
from ..controls import FormCtlSubmitButton
from ..controls import FormCtlTextField
from ..controls import FormCtlTimeField
from ..controls.database import FormCtlDbCheckBox
from ..controls.database import FormCtlDbComboBox
from ..controls.database import FormCtlDbCurrencyField
from ..controls.database import FormCtlDbDateField
from ..controls.database import FormCtlDbFormattedField
from ..controls.database import FormCtlDbListBox
from ..controls.database import FormCtlDbNumericField
from ..controls.database import FormCtlDbPatternField
from ..controls.database import FormCtlDbRadioButton
from ..controls.database import FormCtlDbTextField
from ..controls.database import FormCtlDbTimeField

if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from com.sun.star.lang import EventObject
    from ooodev.units import UnitT
    from ooodev.utils.type_var import PathOrStr
    from ..controls.form_ctl_base import FormCtlBase
# endregion Other Imports

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from com.sun.star.form.component import Form
    from ooodev.units import UnitT

from ooodev.proto.component_proto import ComponentT


class FormPartial:
    """
    Method for adding controls and other form elements to a form.
    """

    def __init__(self, owner: ComponentT, draw_page: XDrawPage, component: Form) -> None:
        """
        Constructor

        Args:
            owner (ComponentT): Class that owns this component.
            draw_page (XDrawPage): Draw Page
            component (Form): Form component
        """
        assert mInfo.Info.support_service(
            component, "com.sun.star.form.component.Form"
        ), "component must support com.sun.star.form.component.Form service"
        self.__owner = owner
        self.__component = component
        self.__draw_page = draw_page
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
        return mForms.Forms.insert_control_button(
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
            parent_form (XNameContainer, optional): Parent form in which to add control.
            styles (Iterable[StyleT], optional): One or more styles to apply to the control shape.

        Returns:
            FormCtlCheckBox: Checkbox Control

        .. versionadded:: 0.14.0
        """
        return mForms.Forms.insert_control_check_box(
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

    # endregion Insert Control Methods
