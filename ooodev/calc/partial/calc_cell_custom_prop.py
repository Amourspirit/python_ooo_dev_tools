from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple
import contextlib
import uno
import unohelper
from com.sun.star.container import XContainerListener
from com.sun.star.drawing import XControlShape
from com.sun.star.form import XForm
from com.sun.star.lang import XComponent
from com.sun.star.awt import XControlModel

# from com.sun.star.xml import AttributeData
from ooo.dyn.drawing.text_vertical_adjust import TextVerticalAdjust
from ooo.dyn.beans.property_attribute import PropertyAttributeEnum
from ooo.dyn.awt.size import Size
from ooo.dyn.awt.point import Point

from ooodev.calc.calc_cell import CalcCell
from ooodev.form.controls.form_ctl_hidden import FormCtlHidden
from ooodev.utils import gen_util as gUtil
from ooodev.utils import props as mProps
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.form.forms import Forms

if TYPE_CHECKING:
    from com.sun.star.form.component import Form
    from com.sun.star.container import ContainerEvent
    from com.sun.star.container import XContainer
    from com.sun.star.drawing import ControlShape
    from com.sun.star.lang import EventObject

    # from com.sun.star.xml import AttributeContainer
    from ooodev.loader.inst.lo_inst import LoInst

# It is fine to add shape UserDefinedAttributes and they will persist when the doc is saved.
# However, the hidden control id can be made part of the shape name.
# This reduces the need to store the id in the shape UserDefinedAttributes.
# This has the added benefit of less clutter in the document.
#
# Because the shape that is added to the cell is relative to the cell then when the cell moves the shape moves with it.
# The shapes anchor is a reference to the cell. When the cell is moved the shape Anchor is updated.
# The cell shape contains a reference to the hidden control that holds the custom properties.

# Known Issues:
# 1. When a cell is copied and pasted the custom properties are not copied.
# However, the shape is copied in the same cell.
# May end up with duplicate shapes something like
# <table:table-cell office:value-type="float" office:value="11" calcext:value-type="float">
# <text:p>11</text:p>
# <draw:control table:end-cell-address="Sheet1.A4" table:end-x="1mm" table:end-y="1mm" draw:z-index="1" draw:name="_cprop_idbldsp50aj59qel_id" draw:style-name="gr1" draw:text-style-name="P1" drawooo:display="none" svg:width="1mm" svg:height="1mm" svg:x="0mm" svg:y="0mm"/>
# <draw:control table:end-cell-address="Sheet1.A4" table:end-x="1mm" table:end-y="1mm" draw:z-index="5" draw:name="_cprop_idbldsp50aj59qel_id 1" draw:style-name="gr1" draw:text-style-name="P1" drawooo:display="none" svg:width="1mm" svg:height="1mm" svg:x="0mm" svg:y="0mm"/>
# </table:table-cell>
# This does not hinder custom properties being added to the copied cell. A new shape will be created for copied cells.
# For good practice these artifacts sound be removed. This is done by the CalcCellCustomProp when calls _find_shape_by_cell_row_col() method.
#
# 2. When a cell is deleted the hidden control for the custom properties are not deleted.
# This is because the hidden control is not part of the cell but part of the form.
# This has no ill effect on the document. To remove all the hidden controls at one the CellCustomProperties Form could be deleted. Although not recommended.
# Removing the form would not remove the shapes.


class CalcCellCustomProp:
    """A partial class for Calc Cell custom properties."""

    class ContainerListener(unohelper.Base, XContainerListener):

        def __init__(
            self, form_name: str, cp: CalcCellCustomProp, lo_inst: LoInst, subscriber: XContainer | None = None
        ) -> None:
            super().__init__()
            self._form_name = form_name
            self._cp = cp
            self._lo_inst = lo_inst
            if subscriber:
                subscriber.addContainerListener(self)

        def is_element_monitored_form(self, element: Any) -> bool:
            form = self.lo_inst.qi(XForm, element)
            if form is None:
                return False
            return form.Name == self._form_name  # type: ignore

        def reset(self) -> None:
            self._cp._reset()

        # region XContainerListener
        def elementInserted(self, event: ContainerEvent) -> None:
            """
            Event is invoked when a container has inserted an element.
            """
            # replaced element should be a form
            if self.is_element_monitored_form(event.Element):
                self.reset()

        def elementRemoved(self, event: ContainerEvent) -> None:
            """
            Event is invoked when a container has removed an element.
            """
            if self.is_element_monitored_form(event.Element):
                self.reset()

        def elementReplaced(self, event: ContainerEvent) -> None:
            """
            Event is invoked when a container has replaced an element.
            """
            if self.is_element_monitored_form(event.ReplacedElement):
                self.reset()

        def disposing(self, event: EventObject) -> None:
            """
            Gets called when the broadcaster is about to be disposed.

            All listeners and all other objects, which reference the broadcaster
            should release the reference to the source. No method should be invoked
            anymore on this object ( including ``XComponent.removeEventListener()`` ).

            This method is called for every listener registration of derived listener
            interfaced, not only for registrations at ``XComponent``.
            """
            # from com.sun.star.lang.XEventListener
            self.reset()

        @property
        def lo_inst(self) -> LoInst:
            return self._lo_inst

    def __init__(self, cell: CalcCell, form_name: str = "CellCustomProperties") -> None:
        self._cell = cell
        self._shape_prefix = "_cprop_"
        self._shape_suffix = "_id"  # suffix is important for ensure shape duplicates are removed.
        self._draw_page = self._cell.calc_sheet.draw_page
        self._forbidden_keys = set(("HiddenValue", "Name", "ClassId", "Tag"))
        self._form_name = form_name
        self._attribute_name = "CustomPropertiesId"
        self._ctl_name = None
        self._row = self._cell.cell_obj.row - 1
        self._col = self._cell.cell_obj.col_obj.index
        self._cache = {}
        self._allow_ensure = True
        self._container_listener: CalcCellCustomProp.ContainerListener
        self._container_listener = CalcCellCustomProp.ContainerListener(
            form_name=self._form_name, cp=self, lo_inst=self._cell.lo_inst, subscriber=self._draw_page.forms.component
        )

    # region Manage Cell Shape
    def _get_control_id(self) -> str:
        key = "control_id"
        if key in self._cache:
            return self._cache[key]
        shape = self._find_shape_by_cell_row_col(self._row, self._col)
        if shape is None:
            s = self._add_shape_to_cell()
            return s
        prefix_len = len(self._shape_prefix)
        suffix_len = len(self._shape_suffix)
        # Name is in format of _cprop_idofhsvtcky1hgom_id
        s = cast(str, shape.Name[prefix_len:])  # type: ignore
        s = s[:-suffix_len]

        # container = cast("AttributeContainer", shape.UserDefinedAttributes)  # type: ignore
        # s = container.getByName(self._attribute_name).Value
        self._cache[key] = s
        return s

    def _find_shape_by_cell_row_col(self, row: int, col: int) -> XControlShape | None:
        key = f"shape_{row}_{col}"
        if key in self._cache:
            return self._cache[key]
        comp = self._draw_page.component
        cleanup = []
        result = None
        for shape in comp:  # type: ignore
            if not shape.supportsService("com.sun.star.drawing.ControlShape"):
                continue

            anchor = shape.Anchor
            if anchor is None:
                continue
            if anchor.getImplementationName() != "ScCellObj":
                continue

            if not shape.Name.startswith(self._shape_prefix):
                continue

            cell_address = anchor.CellAddress
            if cell_address.Row == row and cell_address.Column == col:
                if shape.Name.endswith(self._shape_suffix):
                    self._cache[key] = shape
                    # shape.Anchor = self._cell.component  # type: ignore
                    result = shape
                else:
                    cleanup.append(shape)

        if cleanup:
            for shape in cleanup:
                self._draw_page.remove(shape)
        return result

    def _create_shape(self) -> XControlShape:
        shape = cast(
            "ControlShape",
            self._cell.lo_inst.create_instance_msf(XControlShape, "com.sun.star.drawing.ControlShape", raise_err=True),
        )
        # Setting shape.Visible does not work here. It does work after shape has been added to the draw page.
        # shape.setPropertyValue("Visible", False)
        # shape.Visible = False

        return shape

    def _add_shape_to_cell(self) -> str:
        # There is a strange issue, maybe a bug, when the shape is added the the cell the sheet must be activated.
        # If the sheet is not active when the shape is added then everything seems to work fine until you try to save the document.
        # The document will hang and not save.
        # Testing is showing that by activate the sheet before adding the shape the issue is resolved.
        # This also works in headless mode.
        # Once a cell has a shape for the custom properties it will not be added again and this is no longer an issue.
        active_sheet = self._cell.calc_doc.get_active_sheet()
        activated = False
        if active_sheet != self._cell.calc_sheet:
            activated = True
            self._cell.calc_sheet.set_active()
        shape = cast("ControlShape", self._create_shape())
        # the name is not critical as long as it starts with the prefix
        str_id = "id" + gUtil.Util.generate_random_alpha_numeric(14).lower()
        shape.Name = f"{self._shape_prefix}{str_id}{self._shape_suffix}"  # type: ignore
        # attempting to  set shape.ZOrder does not work here.

        # container = cast("AttributeContainer", shape.UserDefinedAttributes)  # type: ignore

        # ad = AttributeData()  # only CDATA which is the default for Type seems to work
        # ad.Type = "CDATA"
        # ad.Value = str_id
        # container.insertByName(self._attribute_name, ad)
        # shape.UserDefinedAttributes = container  # type: ignore

        # FixedText
        model = self._cell.lo_inst.create_instance_mcf(
            XControlModel, f"com.sun.star.form.component.FixedText", raise_err=True
        )
        model.Name = f"Model_{shape.Name}"  # type: ignore
        model.Label = ""  # type: ignore
        shape.setControl(model)

        model = shape.getControl()
        form = self._get_form()
        name = Forms.create_name(form, "Control")
        form.insertByName(name, model)

        model_key_val = {"Printable": False, "EnableVisible": False, "Enabled": False}
        for key, val in model_key_val.items():
            if hasattr(model, key):
                setattr(model, key, val)

        self._draw_page.add(shape)
        # shape.ResizeWithCell = True  # type: ignore
        # shape.MoveProtect = False

        # note setting visible to true here cause the document to hang when be saved. This is a critical failure.
        # The Control can be set to invisible but the shape must remain visible.
        shape_key_val = {
            "Decorative": False,
            "TextVerticalAdjust": TextVerticalAdjust.CENTER,
            "HoriOrient": 0,
            "MoveProtect": False,
            "Printable": False,
            "ResizeWithCell": True,
            "SizeProtect": False,
            "Visible": True,
        }
        for key, val in shape_key_val.items():
            if hasattr(shape, key):
                setattr(shape, key, val)

        shape.Anchor = self._cell.component  # type: ignore

        shape.setSize(Size(1, 1))

        with contextlib.suppress(Exception):
            control_access = self._cell.calc_doc.get_control_access()
            ctl = control_access.getControl(model)
            if not ctl.isDesignMode():
                ctl.setDesignMode(True)
                ctl.setVisible(False)  # type: ignore
            ctl.setDesignMode(False)

        if activated:
            active_sheet.set_active()
        return str_id

    def _get_pos_size(self) -> Tuple[int, int, int, int]:
        ps = self._cell.component.Position
        size = self._cell.component.Size
        return (ps.X, ps.Y, size.Width, size.Height)

    # endregion Manage Cell Shape

    # region Manage Hidden Control

    def _get_form(self) -> Form:

        key = self._form_name
        if key in self._cache:
            return self._cache[key]
        forms = self._draw_page.forms.component
        if len(forms) == 0:  # type: ignore
            # insert a default form1.
            # The reason for this is many users many working in forms[0].
            # This way there will be a from to work with that is not for properties.
            # This is not critical but it is a good practice.
            # If the user deletes Forms[0] it will not wipe the property forms.
            # Also if the user draws control on the spreadsheet or other document it will use this form.
            frm = cast(
                "Form",
                self._cell.lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True),
            )
            frm.Name = "Form1"
            forms.insertByName("Form1", frm)

        if forms.hasByName(key):
            frm = forms.getByName(key)
        else:
            frm = cast(
                "Form",
                self._cell.lo_inst.create_instance_mcf(XForm, "stardiv.one.form.component.Form", raise_err=True),
            )
            frm.Name = key
            forms.insertByName(key, frm)
        self._cache[key] = frm
        return frm

    def _get_hidden_control(self) -> FormCtlHidden:
        ctl_name = self._get_control_id()
        key = f"hidden_ctl_{ctl_name}"
        if key in self._cache:
            return cast(FormCtlHidden, self._cache[key])
        frm = self._get_form()
        if not frm.hasByName(ctl_name):
            comp = self._cell.lo_inst.create_instance_mcf(
                XComponent, "com.sun.star.form.component.HiddenControl", raise_err=True
            )
            comp.HiddenValue = self.__class__.__qualname__  # type: ignore
            frm.insertByName(ctl_name, comp)

        ctl = FormCtlHidden(frm.getByName(ctl_name), self._cell.lo_inst)
        self._cache[key] = ctl
        return ctl

    # endregion Manage Hidden Control

    # region Property Access

    def get_custom_property(self, name: str, default: Any = NULL_OBJ) -> Any:
        """
        Gets a custom property.

        Args:
            name (str): The name of the property.
            default (Any, optional): The default value to return if the property does not exist.

        Raises:
            AttributeError: If the property is not found.

        Returns:
            Any: The value of the property.
        """
        ctl = self._get_hidden_control()
        info = ctl.get_property_set_info()
        if info.hasPropertyByName(name):
            return ctl.get_property(name)
        if default is not NULL_OBJ:
            return default
        raise AttributeError(f"Property '{name}' not found.")

    def get_custom_properties(self) -> DotDict:
        """
        Gets custom properties.

        Returns:
            DotDict: custom properties.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        ctl = self._get_hidden_control()
        props = ctl.get_property_values()
        lst = []
        for prop in props:
            if prop.Name not in self._forbidden_keys:
                lst.append(prop)
        return mProps.Props.props_to_dot_dict(lst)

    def has_custom_property(self, name: str) -> bool:
        """
        Gets if a custom property exists.

        Args:
            name (str): The name of the property to check.

        Returns:
            bool: ``True`` if the property exists, otherwise ``False``.
        """
        key = self._form_name
        if key not in self._cache:
            # form has not been loaded yet, It may not exist
            if not self._draw_page.forms.has_by_name(self._form_name):
                # if there is no form there is no properties for any cell yet.
                return False

        key = "control_id"
        if key not in self._cache:
            # shape has not been loaded
            shape = self._find_shape_by_cell_row_col(self._row, self._col)
            # if the cell has no shape it has no properties
            if shape is None:
                return False

        ctl = self._get_hidden_control()
        info = ctl.get_property_set_info()
        return info.hasPropertyByName(name)

    def set_custom_property(self, name: str, value: Any):
        """
        Sets a custom property.

        Args:
            name (str): The name of the property.
            value (Any): The value of the property.

        Raises:
            AttributeError: If the property is a forbidden key.
        """
        if name in self._forbidden_keys:
            raise AttributeError(f"Property '{name}' is forbidden. Forbidden keys: {self._forbidden_keys}")
        ctl = self._get_hidden_control()
        info = ctl.get_property_set_info()
        if info.hasPropertyByName(name):
            ctl.remove_property(name)

        ctl.add_property(name, PropertyAttributeEnum.REMOVABLE, value)

    def set_custom_properties(self, properties: DotDict) -> None:
        """
        Sets custom properties.

        Args:
            properties (DotDict): custom properties to set.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        for name, value in properties.items():
            self.set_custom_property(name, value)

    def remove_custom_property(self, name: str) -> None:
        """
        Removes a custom property.

        Args:
            name (str): The name of the property to remove.

        Raises:
            AttributeError: If the property is a forbidden key.

        Returns:
            None:
        """
        if name in self._forbidden_keys:
            raise AttributeError(f"Property '{name}' is forbidden. Forbidden keys: {self._forbidden_keys}")
        ctl = self._get_hidden_control()
        info = ctl.get_property_set_info()
        if info.hasPropertyByName(name):
            ctl.remove_property(name)

    def remove_custom_properties(self) -> None:
        """
        Removes all custom properties.

        Returns:
            None:
        """
        # remove hidden control
        # remove form if it is empty
        # remove shape
        forms = self._draw_page.forms
        if forms.has_by_name(self._form_name):
            ctl_id = self._get_control_id()
            form = forms[self._form_name]
            form.remove_by_name(ctl_id)
        if not form.has_elements():
            forms.remove_by_name(self._form_name)
            form = None
        shape = self._find_shape_by_cell_row_col(self._row, self._col)
        if shape:
            self._draw_page.remove(shape)
            shape = None

    # endregion Property Access

    def _reset(self) -> None:
        self._cache.clear()

    def __del__(self) -> None:
        with contextlib.suppress(Exception):
            comp = self._draw_page.forms.component
            if self._container_listener and comp:
                comp.removeContainerListener(self._container_listener)
            self._container_listener = None  # type: ignore
