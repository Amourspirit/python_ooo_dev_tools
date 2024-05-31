from __future__ import annotations
from typing import Any, cast, Dict, List, TYPE_CHECKING, Tuple
import contextlib
import uno
import unohelper
from com.sun.star.container import XContainerListener
from com.sun.star.drawing import XControlShape
from com.sun.star.form import XForm
from com.sun.star.lang import XComponent

from ooo.dyn.drawing.text_vertical_adjust import TextVerticalAdjust
from ooo.dyn.beans.property_attribute import PropertyAttributeEnum
from ooo.dyn.awt.size import Size

from ooodev.calc.calc_cell import CalcCell
from ooodev.calc.cell.custom_prop_base import CustomPropBase
from ooodev.form.controls.form_ctl_hidden import FormCtlHidden
from ooodev.utils import gen_util as gUtil
from ooodev.utils import props as mProps
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.utils.helper.dot_dict import DotDict

if TYPE_CHECKING:
    from com.sun.star.container import ContainerEvent
    from com.sun.star.container import XContainer
    from com.sun.star.drawing import ControlShape
    from com.sun.star.lang import EventObject
    from ooodev.loader.inst.lo_inst import LoInst

# Because the shape that is added to the cell is relative to the cell then when the cell moves the shape moves with it.
# The shapes anchor is a reference to the cell. When the cell is moved the shape Anchor is updated.
# The cell shape contains a reference to the hidden control that holds the custom properties.

# Known Issues:
# 1. When a cell is copied and pasted the custom properties are not copied. This is by design.
# However, when cell is copied it will copy the shape that is in the cell. The duplicated shape will be automatically removed then the next call to original cell or the copied cell for custom properties.
#
# 2. When a cell is deleted the hidden control for the custom properties are not deleted.
# This is because the hidden control is not part of the cell but part of the form.
# This has no ill effect on the document. To remove all the hidden controls at one the CellCustomProperties Form could be deleted. Although not recommended.
# Removing the form would not remove the shapes. It would be possible for a developer to monitor the sheet and manually remove the hidden controls.
# This is not really an issue because the hidden controls are not visible and do not effect the document.
# See Also: https://ask.libreoffice.org/t/how-to-detect-cell-delete-event/106250
# Also There is a ooodev.calc.cell.custom_prop_clean.CustomPropClean that can be used to clean up the hidden controls if needed.
# this could also be done on a document saving event or other event if needed


class CustomProp(CustomPropBase):
    """A partial class for Calc Cell custom properties."""

    class ContainerListener(unohelper.Base, XContainerListener):

        def __init__(
            self, form_name: str, cp: CustomProp, lo_inst: LoInst, subscriber: XContainer | None = None
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
            self.reset()

        @property
        def lo_inst(self) -> LoInst:
            return self._lo_inst

    def __init__(self, cell: CalcCell) -> None:
        CustomPropBase.__init__(self, cell.calc_sheet)
        self._cell = cell
        self._forbidden_keys = set(("HiddenValue", "Name", "ClassId", "Tag"))
        self._attribute_name = "CustomPropertiesId"
        self._ctl_name = None
        self._row = self._cell.cell_obj.row - 1
        self._col = self._cell.cell_obj.col_obj.index
        self._container_listener: CustomProp.ContainerListener
        self._container_listener = CustomProp.ContainerListener(
            form_name=self._form_name, cp=self, lo_inst=self._cell.lo_inst, subscriber=self._draw_page.forms.component
        )

    # region Manage Cell Shape

    def _get_control_id(self) -> Tuple[XControlShape, str]:
        key = "control_id"
        if key in self.cache:
            return self.cache[key]
        shape = self._find_shape_by_cell_row_col(self._row, self._col)
        if shape is None:
            shape, s = self._add_shape_to_cell()
            return shape, s
        s = self._get_hidden_control_name_from_shape(shape)

        self.cache[key] = shape, s
        return shape, s

    def _get_shapes_dict(self) -> Dict[str, List[XControlShape]]:

        comp = self.draw_page.component
        shapes = {}
        # find all shapes on the draw page that start with prefix and end with suffix
        for shape in comp:  # type: ignore
            if not shape.supportsService("com.sun.star.drawing.ControlShape"):
                continue
            name = cast(str, shape.Name)
            if name.startswith(self.shape_prefix) and name.endswith(self.shape_suffix):
                if name in shapes:
                    shapes[name].append(shape)
                else:
                    shapes[name] = [shape]
        return shapes

    def _find_shape_by_cell_row_col(self, row: int, col: int) -> XControlShape | None:
        # When a cell has been copied and pasted there will be a duplicate shape in the dest cell.
        # For this reason a small cleanup is done if needed to remove any duplicate shapes.
        # The duplicate shapes will have a higher z-order then the original shape. The higher z-order shapes are removed.
        # Depending on environment there may not be duplicates with the exact shape name. It may be an artifact such as '_cprop_idhdkuy07hizr3eh_id 1'
        # In this case the artifact is removed when the cell custom properties that contains the artifact is accessed.

        key = f"shape_{row}_{col}"
        if key in self.cache:
            return self.cache[key]
        comp = self.draw_page.component

        found_shape = None
        cleanup = []

        for shape in comp:  # type: ignore
            if not shape.supportsService("com.sun.star.drawing.ControlShape"):
                continue

            anchor = shape.Anchor
            if anchor is None:
                continue
            if anchor.getImplementationName() != "ScCellObj":
                continue

            if not shape.Name.startswith(self.shape_prefix):
                continue

            cell_address = anchor.CellAddress
            if cell_address.Row == row and cell_address.Column == col:
                if shape.Name.endswith(self.shape_suffix):
                    if found_shape is None:
                        found_shape = shape
                else:
                    cleanup.append(shape)

        if found_shape is None:
            if cleanup:
                for shape in cleanup:
                    with contextlib.suppress(Exception):
                        self._draw_page.remove(shape)
            cleanup.clear()
            return None

        def get_result(shp: Any) -> Any:
            nonlocal row, col
            anchor = shp.Anchor
            if anchor is None:
                return None
            if anchor.getImplementationName() != "ScCellObj":
                return None
            cell_address = anchor.CellAddress
            if cell_address.Row == row and cell_address.Column == col:
                # the current shape may be the original shape in another cell.
                # That means this cell is a copy.
                return shp
            return None

        result = None

        shapes_dict = self._get_shapes_dict()
        shapes = shapes_dict[found_shape.Name]
        if len(shapes) == 1:
            result = found_shape
        else:
            # sort shapes by zorder
            found_shape = None
            result = None
            shapes.sort(key=lambda x: x.ZOrder)  # type: ignore
            # if there is more then one shape then the first shape is the original
            # if cleanup is not empty at this point then it will contain artifacts such as '_cprop_idhdkuy07hizr3eh_id 1' that need to be removed.
            cleanup.extend(shapes[1:])
            result = get_result(shapes[0])

        if cleanup:
            for shape in cleanup:
                with contextlib.suppress(Exception):
                    self.draw_page.remove(shape)
        if result:
            self.cache[key] = result
        return result

    def _create_shape(self) -> XControlShape:
        shape = cast(
            "ControlShape",
            self.sheet.lo_inst.create_instance_msf(XControlShape, "com.sun.star.drawing.ControlShape", raise_err=True),
        )
        # Setting shape.Visible does not work here. It does work after shape has been added to the draw page.
        # shape.setPropertyValue("Visible", False)
        # shape.Visible = False

        return shape

    def _add_shape_to_cell(self) -> Tuple[XControlShape, str]:
        # There is a strange issue, maybe a bug, when the shape is added the the cell the sheet must be activated.
        # If the sheet is not active when the shape is added then everything seems to work fine until you try to save the document.
        # The document will hang and not save.
        # Testing is showing that by activate the sheet before adding the shape the issue is resolved.
        # This also works in headless mode.
        # Once a cell has a shape for the custom properties it will not be added again and this is no longer an issue.
        active_sheet = self.sheet.calc_doc.get_active_sheet()
        activated = False
        if active_sheet.name != self.sheet.name:
            activated = True
            self.sheet.calc_sheet.set_active()
        shape = cast("ControlShape", self._create_shape())

        str_id = "id" + gUtil.Util.generate_random_alpha_numeric(14).lower()
        shape.Name = f"{self.shape_prefix}{str_id}{self.shape_suffix}"  # type: ignore

        self.draw_page.add(shape)

        # note setting visible to true here cause the document to hang when be saved. This is a critical failure.
        shape.Anchor = self._cell.component  # type: ignore
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

        shape.setSize(Size(1, 1))

        if activated:
            active_sheet.set_active()
        return shape, str_id

    # endregion Manage Cell Shape

    # region Manage Hidden Control

    def _get_hidden_control(self) -> FormCtlHidden:
        _, ctl_name = self._get_control_id()
        key = f"hidden_ctl_{ctl_name}"
        if key in self.cache:
            return cast(FormCtlHidden, self.cache[key])
        frm = self._get_form()
        if not frm.hasByName(ctl_name):
            comp = self._cell.lo_inst.create_instance_mcf(
                XComponent, "com.sun.star.form.component.HiddenControl", raise_err=True
            )
            comp.HiddenValue = ""  # type: ignore
            frm.insertByName(ctl_name, comp)

        ctl = FormCtlHidden(frm.getByName(ctl_name), self.sheet.lo_inst)
        self.cache[key] = ctl
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
        key = self.form_name
        if key not in self.cache:
            # form has not been loaded yet, It may not exist
            if not self.draw_page.forms.has_by_name(self._form_name):
                # if there is no form there is no properties for any cell yet.
                return False

        key = "control_id"
        if key not in self.cache:
            # shape has not been loaded
            shape = self._find_shape_by_cell_row_col(self._row, self._col)
            # if the cell has no shape it has no properties
            if shape is None:
                return False

        ctl = self._get_hidden_control()
        info = ctl.get_property_set_info()
        return info.hasPropertyByName(name)

    def has_custom_properties(self) -> bool:
        """
        Gets if a custom property exists.

        Args:
            name (str): The name of the property to check.

        Returns:
            bool: ``True`` if the property exists, otherwise ``False``.
        """
        key = self.form_name
        if key not in self.cache:
            # form has not been loaded yet, It may not exist
            if not self.draw_page.forms.has_by_name(self.form_name):
                # if there is no form there is no properties for any cell yet.
                return False

        key = "control_id"
        shape = None
        if key not in self.cache:
            # shape has not been loaded
            shape = self._find_shape_by_cell_row_col(self._row, self._col)
            # if the cell has no shape it has no properties
        if shape is None:
            return False
        hidden = self._get_hidden_control_name_from_shape(shape)
        if hidden is None:
            return False
        props = self.get_custom_properties()
        return len(props) > 0

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
        forms = self.draw_page.forms
        if forms.has_by_name(self.form_name):
            _, ctl_id = self._get_control_id()
            form = forms[self.form_name]
            form.remove_by_name(ctl_id)
        if not form.has_elements():
            forms.remove_by_name(self.form_name)
            form = None
        shape = self._find_shape_by_cell_row_col(self._row, self._col)
        if shape:
            self._draw_page.remove(shape)
            shape = None

    # endregion Property Access

    def __del__(self) -> None:
        with contextlib.suppress(Exception):
            comp = self._draw_page.forms.component
            if self._container_listener and comp:
                comp.removeContainerListener(self._container_listener)
            self._container_listener = None  # type: ignore
        super().__del__()
