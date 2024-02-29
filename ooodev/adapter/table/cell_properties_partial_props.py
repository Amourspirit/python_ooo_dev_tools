"""
This class is a partial class for ``com.sun.star.table.CellProperties`` service.
"""

from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING, Tuple
import uno
from com.sun.star.container import XNameContainer

from ooo.dyn.table.cell_vert_justify2 import CellVertJustify2Enum
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.units.angle100 import Angle100
from ooodev.utils import info as mInfo

from ooodev.adapter.container.name_container_comp import NameContainerComp
from ooodev.adapter.table.border_line_comp import BorderLineComp
from ooodev.adapter.table.border_line2_comp import BorderLine2Comp
from ooodev.adapter.table.table_border_comp import TableBorderComp
from ooodev.adapter.table.table_border2_comp import TableBorder2Comp
from ooodev.adapter.util.cell_protection_comp import CellProtectionComp
from ooodev.adapter.table.shadow_format_comp import ShadowFormatComp
from ooodev.events.events import Events


if TYPE_CHECKING:
    from com.sun.star.table import CellProperties  # service
    from com.sun.star.util import CellProtection  # Struct
    from com.sun.star.table import TableBorder
    from com.sun.star.table import TableBorder2
    from com.sun.star.beans import PropertyValue
    from com.sun.star.table import ShadowFormat
    from com.sun.star.table import BorderLine
    from com.sun.star.table import BorderLine2
    from ooodev.utils.color import Color
    from ooo.lo.table.cell_hori_justify import CellHoriJustify
    from ooo.lo.table.cell_orientation import CellOrientation
    from ooodev.units.angle_t import AngleT
    from ooodev.units.unit_obj import UnitT
    from ooodev.events.args.key_val_args import KeyValArgs


class CellPropertiesPartialProps:
    """
    Partial Class for CellProperties Service.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: CellProperties) -> None:
        """
        Constructor

        Args:
            component (CellProperties): UNO Component that implements ``com.sun.star.table.CellProperties`` service.
        """
        self.__component = component
        self.__event_provider = Events(self)
        self.__props = {}

        def on_comp_struct_changed(src: Any, event_args: KeyValArgs) -> None:
            prop_name = str(event_args.event_data["prop_name"])
            if hasattr(self.__component, prop_name):
                setattr(self.__component, prop_name, event_args.source.component)

        self.__fn_on_comp_struct_changed = on_comp_struct_changed
        # pylint: disable=no-member
        self.__event_provider.subscribe_event(
            "com_sun_star_table_BorderLine_changed", self.__fn_on_comp_struct_changed
        )
        self.__event_provider.subscribe_event(
            "com_sun_star_table_BorderLine2_changed", self.__fn_on_comp_struct_changed
        )
        self.__event_provider.subscribe_event(
            "com_sun_star_util_CellProtection_changed", self.__fn_on_comp_struct_changed
        )
        self.__event_provider.subscribe_event(
            "com_sun_star_table_TableBorder_changed", self.__fn_on_comp_struct_changed
        )
        self.__event_provider.subscribe_event(
            "com_sun_star_table_TableBorder2_changed", self.__fn_on_comp_struct_changed
        )
        self.__event_provider.subscribe_event(
            "com_sun_star_table_ShadowFormat_changed", self.__fn_on_comp_struct_changed
        )

    # region Properties
    @property
    def cell_interop_grab_bag(self) -> Tuple[PropertyValue, ...] | None:
        """
        Gets/Sets Grab bag of cell properties, used as a string-any map for interim interop purposes.

        This property is intentionally not handled by the ODF filter. Any member that should be handled there should be first moved out from this grab bag to a separate property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.CellInteropGrabBag
        return None

    @cell_interop_grab_bag.setter
    def cell_interop_grab_bag(self, value: Tuple[PropertyValue, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.CellInteropGrabBag = value

    @property
    def asian_vertical_mode(self) -> bool | None:
        """
        Gets/Sets Asian character orientation in vertical orientation.

        If the CellProperties.Orientation property is CellOrientation.STACKED, in Asian mode only Asian characters are printed in horizontal orientation instead of all characters. For other values of CellProperties.Orientation, this value is not used.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.AsianVerticalMode
        return None

    @asian_vertical_mode.setter
    def asian_vertical_mode(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.AsianVerticalMode = value

    @property
    def bottom_border(self) -> BorderLineComp:
        """
        Gets/Sets a description of the bottom border line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLineComp`` object.

        Returns:
            BorderLineComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "BottomBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLineComp(self.__component.BottomBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLineComp, prop)

    @bottom_border.setter
    def bottom_border(self, value: BorderLine | BorderLineComp) -> None:
        key = "BottomBorder"
        if mInfo.Info.is_instance(value, BorderLineComp):
            self.__component.BottomBorder = value.component
        else:
            self.__component.BottomBorder = cast("BorderLine", value)
        self.__props[key] = BorderLineComp(self.__component.BottomBorder, key, self.__event_provider)

    @property
    def bottom_border2(self) -> BorderLine2Comp | None:
        """
        Gets/Sets a description of the bottom border line of each cell.

        Preferred over ``bottom_border``.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        **optional**

        Returns:
            BorderLine2Comp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "BottomBorder2"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2Comp(self.__component.BottomBorder2, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2Comp, prop)

    @bottom_border2.setter
    def bottom_border2(self, value: BorderLine2 | BorderLine2Comp) -> None:
        key = "BottomBorder2"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2Comp):
            self.__component.BottomBorder2 = value.component
        else:
            self.__component.BottomBorder2 = cast("BorderLine2", value)
        self.__props[key] = BorderLine2Comp(self.__component.BottomBorder2, key, self.__event_provider)

    @property
    def cell_back_color(self) -> Color:
        """
        Gets/Sets the cell background color.

        Returns:
            ~ooo.dyn.utils.color.Color: Returns Color.
        """
        return self.__component.CellBackColor  # type: ignore

    @cell_back_color.setter
    def cell_back_color(self, value: Color) -> None:
        self.__component.CellBackColor = value  # type: ignore

    @property
    def cell_protection(self) -> CellProtectionComp:
        """
        Gets/Sets a description of the cell protection.

        Cell protection is active only if the sheet is protected.
        """
        key = "CellProtection"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = CellProtectionComp(self.__component.CellProtection, key, self.__event_provider)
            self.__props[key] = prop
        return cast(CellProtectionComp, prop)

    @cell_protection.setter
    def cell_protection(self, value: CellProtection | CellProtectionComp) -> None:
        key = "CellProtection"
        if mInfo.Info.is_instance(value, CellProtectionComp):
            self.__component.CellProtection = value.component
        else:
            self.__component.CellProtection = cast("CellProtection", value)
        self.__props[key] = CellProtectionComp(self.__component.CellProtection, key, self.__event_provider)

    @property
    def cell_style(self) -> str:
        """
        Gets/Sets the name of the style of the cell.
        """
        return self.__component.CellStyle

    @cell_style.setter
    def cell_style(self, value: str) -> None:
        self.__component.CellStyle = value

    @property
    def diagonal_bltr(self) -> BorderLineComp | None:
        """
        Gets/Sets a description of the bottom left to top right diagonal line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLineComp`` object.

        **optional**

        Returns:
            BorderLineComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "DiagonalBLTR"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLineComp(self.__component.DiagonalBLTR, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLineComp, prop)

    @diagonal_bltr.setter
    def diagonal_bltr(self, value: BorderLine | BorderLineComp) -> None:
        key = "DiagonalBLTR"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLineComp):
            self.__component.DiagonalBLTR = value.component
        else:
            self.__component.DiagonalBLTR = cast("BorderLine", value)
        self.__props[key] = BorderLineComp(self.__component.DiagonalBLTR, key, self.__event_provider)

    @property
    def diagonal_bltr2(self) -> BorderLine2Comp | None:
        """
        Gets/Sets a description of the bottom left to top right diagonal line of each cell.

        Preferred over ``diagonal_bltr``.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        **optional**

        Returns:
            BorderLine2Comp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "DiagonalBLTR2"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2Comp(self.__component.DiagonalBLTR2, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2Comp, prop)

    @diagonal_bltr2.setter
    def diagonal_bltr2(self, value: BorderLine2 | BorderLine2Comp) -> None:
        key = "DiagonalBLTR2"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2Comp):
            self.__component.DiagonalBLTR2 = value.component
        else:
            self.__component.DiagonalBLTR2 = cast("BorderLine2", value)
        self.__props[key] = BorderLine2Comp(self.__component.DiagonalBLTR2, key, self.__event_provider)

    @property
    def diagonal_tlbr(self) -> BorderLineComp | None:
        """
        Gets/Sets a description of the top left to bottom right diagonal line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLineComp`` object.

        **optional**

        Returns:
            BorderLineComp: Returns BorderLine.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "DiagonalTLBR"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLineComp(self.__component.DiagonalTLBR, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLineComp, prop)

    @diagonal_tlbr.setter
    def diagonal_tlbr(self, value: BorderLine | BorderLineComp) -> None:
        key = "DiagonalTLBR"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLineComp):
            self.__component.DiagonalTLBR = value.component
        else:
            self.__component.DiagonalTLBR = cast("BorderLine", value)
        self.__props[key] = BorderLineComp(self.__component.DiagonalTLBR, key, self.__event_provider)

    @property
    def diagonal_tlbr2(self) -> BorderLine2Comp | None:
        """
        contains a description of the top left to bottom right diagonal line of each cell.

        Preferred over ``diagonal_tlbr``.

        **optional**

        Returns:
            BorderLine2Comp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "DiagonalTLBR2"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2Comp(self.__component.DiagonalTLBR2, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2Comp, prop)

    @diagonal_tlbr2.setter
    def diagonal_tlbr2(self, value: BorderLine2 | BorderLine2Comp) -> None:
        key = "DiagonalTLBR2"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2Comp):
            self.__component.DiagonalTLBR2 = value.component
        else:
            self.__component.DiagonalTLBR2 = cast("BorderLine2", value)
        self.__props[key] = BorderLine2Comp(self.__component.DiagonalTLBR2, key, self.__event_provider)

    @property
    def hori_justify(self) -> CellHoriJustify:
        """
        Gets/Sets the horizontal alignment of the cell contents.

        Returns:
            CellHoriJustify: Returns Horizontal Justify.

        Hint:
            - ``CellHoriJustifyProto`` can be imported from ``ooo.dyn.table.cell_hori_justify``
        """
        return self.__component.HoriJustify  # type: ignore

    @hori_justify.setter
    def hori_justify(self, value: CellHoriJustify) -> None:
        self.__component.HoriJustify = value  # type: ignore

    @property
    def is_cell_background_transparent(self) -> bool:
        """
        Gets/Sets - is ``True``, if the cell background is transparent.

        In this case the ``CellProperties.CellBackColor`` value is not used.
        """
        return self.__component.IsCellBackgroundTransparent

    @is_cell_background_transparent.setter
    def is_cell_background_transparent(self, value: bool) -> None:
        self.__component.IsCellBackgroundTransparent = value

    @property
    def is_text_wrapped(self) -> bool:
        """
        Gets/Sets - is ``True``, if text in the cells will be wrapped automatically at the right border.
        """
        return self.__component.IsTextWrapped

    @is_text_wrapped.setter
    def is_text_wrapped(self, value: bool) -> None:
        self.__component.IsTextWrapped = value

    @property
    def left_border(self) -> BorderLineComp:
        """
        Gets/Sets a description of the left border line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLineComp`` object.

        Returns:
            BorderLineComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "LeftBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLineComp(self.__component.LeftBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLineComp, prop)

    @left_border.setter
    def left_border(self, value: BorderLine | BorderLineComp) -> None:
        key = "LeftBorder"
        if mInfo.Info.is_instance(value, BorderLineComp):
            self.__component.LeftBorder = value.component
        else:
            self.__component.LeftBorder = cast("BorderLine", value)
        self.__props[key] = BorderLineComp(self.__component.LeftBorder, key, self.__event_provider)

    @property
    def left_border2(self) -> BorderLine2Comp | None:
        """
        Gets/Sets a description of the left border line of each cell.

        Preferred over ``left_border``.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        **optional**

        Returns:
            BorderLine2Comp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "LeftBorder2"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2Comp(self.__component.LeftBorder2, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2Comp, prop)

    @left_border2.setter
    def left_border2(self, value: BorderLine2 | BorderLine2Comp) -> None:
        key = "LeftBorder2"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2Comp):
            self.__component.LeftBorder2 = value.component
        else:
            self.__component.LeftBorder2 = cast("BorderLine2", value)
        self.__props[key] = BorderLine2Comp(self.__component.LeftBorder2, key, self.__event_provider)

    @property
    def number_format(self) -> int:
        """
        Gets/Sets the index of the number format that is used in the cells.

        The proper value can be determined by using the ``com.sun.star.util.NumberFormatter`` interface of the document.
        """
        return self.__component.NumberFormat

    @number_format.setter
    def number_format(self, value: int) -> None:
        self.__component.NumberFormat = value

    @property
    def orientation(self) -> CellOrientation:
        """
        Gets/Sets the orientation of the cell contents.

        If the ``CellProperties.RotateAngle`` property is non-zero, this value is not used.

        Returns:
            CellOrientation: Returns Cell Orientation.

        Hint:
            - ``CellOrientation`` can be imported from ``ooo.dyn.table.cell_orientation``
        """
        return self.__component.Orientation  # type: ignore

    @orientation.setter
    def orientation(self, value: CellOrientation) -> None:
        self.__component.Orientation = value  # type: ignore

    @property
    def para_indent(self) -> UnitMM100:
        """
        Gets/Sets the indentation of the cell contents (in ``1/100 mm``).

        The value can be set with an integer (in ``1/100 mm``) or an ``UnitT`` object.

        Returns:
            UnitMM100: Returns UnitMM100 object (unit ``1/100 mm``).
        """
        return UnitMM100(self.__component.ParaIndent)

    @para_indent.setter
    def para_indent(self, value: int | UnitT) -> None:
        val = UnitMM100.from_unit_val(value)
        self.__component.ParaIndent = val.value

    @property
    def right_border(self) -> BorderLineComp:
        """
        Gets/Sets a description of the right border line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLineComp`` object.

        Returns:
            BorderLineComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "RightBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLineComp(self.__component.RightBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLineComp, prop)

    @right_border.setter
    def right_border(self, value: BorderLine | BorderLineComp) -> None:
        key = "RightBorder"
        if mInfo.Info.is_instance(value, BorderLineComp):
            self.__component.RightBorder = value.component
        else:
            self.__component.RightBorder = cast("BorderLine", value)
        self.__props[key] = BorderLineComp(self.__component.RightBorder, key, self.__event_provider)

    @property
    def right_border2(self) -> BorderLine2Comp | None:
        """
        Gets/Sets a description of the right border line of each cell.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        Preferred over ``right_border``.

        **optional**

        Returns:
            BorderLine2Comp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "RightBorder2"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2Comp(self.__component.RightBorder2, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2Comp, prop)

    @right_border2.setter
    def right_border2(self, value: BorderLine2 | BorderLine2Comp) -> None:
        key = "RightBorder2"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2Comp):
            self.__component.RightBorder2 = value.component
        else:
            self.__component.RightBorder2 = cast("BorderLine2", value)
        self.__props[key] = BorderLine2Comp(self.__component.RightBorder2, key, self.__event_provider)

    @property
    def rotate_angle(self) -> Angle100:
        """
        Gets/Sets how much the content of cells is rotated (in ``1/100`` degrees).

        The value can be set with an integer (in ``1/100`` degrees) or an ``AngleT`` object.
        """
        return Angle100(self.__component.RotateAngle)

    @rotate_angle.setter
    def rotate_angle(self, value: int | AngleT) -> None:
        val = Angle100.from_unit_val(value)
        self.__component.RotateAngle = val.value

    @property
    def rotate_reference(self) -> CellVertJustify2Enum:
        """
        Gets/Sets at which edge rotated cells are aligned.

        Returns:
            CellVertJustify2Enum: Returns CellVertJustify2Enum.

        Hint:
            - ``CellVertJustify2Enum`` can be imported from ``ooo.dyn.table.cell_vert_justify2``
        """
        return CellVertJustify2Enum(self.__component.RotateReference)  # type: ignore

    @rotate_reference.setter
    def rotate_reference(self, value: int | CellVertJustify2Enum) -> None:
        val = CellVertJustify2Enum(value)
        self.__component.RotateReference = val.value

    @property
    def shadow_format(self) -> ShadowFormatComp:
        """
        Gets/Sets a description of the shadow.
        """
        key = "ShadowFormat"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = ShadowFormatComp(self.__component.ShadowFormat, key, self.__event_provider)
            self.__props[key] = prop
        return cast(ShadowFormatComp, prop)

    @shadow_format.setter
    def shadow_format(self, value: ShadowFormat | ShadowFormatComp) -> None:
        key = "ShadowFormat"
        if mInfo.Info.is_instance(value, ShadowFormatComp):
            self.__component.ShadowFormat = value.component
        else:
            self.__component.ShadowFormat = cast("ShadowFormat", value)
        self.__props[key] = ShadowFormatComp(self.__component.ShadowFormat, key, self.__event_provider)

    @property
    def shrink_to_fit(self) -> bool | None:
        """
        Gets/Sets - is ``True``, if the cell content will be shrunk to fit in the cell.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ShrinkToFit
        return None

    @shrink_to_fit.setter
    def shrink_to_fit(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ShrinkToFit = value

    @property
    def table_border(self) -> TableBorderComp:
        """
        contains a description of the cell or cell range border.

        If used with a cell range, the top, left, right, and bottom lines are at the edges of the entire range, not at the edges of the individual cell.

        Returns:
            TableBorderComp: Table Border.
        """
        key = "TableBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = TableBorderComp(self.__component.TableBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(TableBorderComp, prop)

    @table_border.setter
    def table_border(self, value: TableBorder | TableBorderComp) -> None:
        key = "TableBorder"
        if mInfo.Info.is_instance(value, TableBorderComp):
            self.__component.TableBorder = value.component
        else:
            self.__component.TableBorder = cast("TableBorder", value)
        self.__props[key] = TableBorderComp(self.__component.TableBorder, key, self.__event_provider)

    @property
    def table_border2(self) -> TableBorder2Comp | None:
        """
        Gets/Seta a description of the cell or cell range border.

        Preferred over ``table_border``.

        If used with a cell range, the top, left, right, and bottom lines are at the edges of the entire range, not at the edges of the individual cell.

        **optional**
        """
        key = "TableBorder2"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = TableBorder2Comp(self.__component.TableBorder2, key, self.__event_provider)
            self.__props[key] = prop
        return cast(TableBorder2Comp, prop)

    @table_border2.setter
    def table_border2(self, value: TableBorder2 | TableBorder2Comp) -> None:
        key = "TableBorder2"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, TableBorder2Comp):
            self.__component.TableBorder2 = value.component
        else:
            self.__component.TableBorder2 = cast("TableBorder2", value)
        self.__props[key] = TableBorder2Comp(self.__component.TableBorder2, key, self.__event_provider)

    @property
    def top_border(self) -> BorderLineComp:
        """
        Gets/Sets a description of the top border line of each cell.

        Setting value can be done with a ``BorderLine`` or ``BorderLineComp`` object.

        Returns:
            BorderLineComp: Returns Border Line.

        Hint:
            - ``BorderLine`` can be imported from ``ooo.dyn.table.border_line``
        """
        key = "TopBorder"
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLineComp(self.__component.TopBorder, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLineComp, prop)

    @top_border.setter
    def top_border(self, value: BorderLine | BorderLineComp) -> None:
        key = "TopBorder"
        if mInfo.Info.is_instance(value, BorderLineComp):
            self.__component.TopBorder = value.component
        else:
            self.__component.TopBorder = cast("BorderLine", value)
        self.__props[key] = BorderLineComp(self.__component.TopBorder, key, self.__event_provider)

    @property
    def top_border2(self) -> BorderLine2Comp | None:
        """
        Gets/Sets a description of the top border line of each cell.

        Preferred over ``top_border``.

        Setting value can be done with a ``BorderLine2`` or ``BorderLine2Comp`` object.

        **optional**

        Returns:
            BorderLine2Comp | None: Returns BorderLine2 or None if not supported.

        Hint:
            - ``BorderLine2`` can be imported from ``ooo.dyn.table.border_line2``
        """
        key = "TopBorder2"
        if not hasattr(self.__component, key):
            return None
        prop = self.__props.get(key, None)
        if prop is None:
            prop = BorderLine2Comp(self.__component.TopBorder2, key, self.__event_provider)
            self.__props[key] = prop
        return cast(BorderLine2Comp, prop)

    @top_border2.setter
    def top_border2(self, value: BorderLine2 | BorderLine2Comp) -> None:
        key = "TopBorder2"
        if not hasattr(self.__component, key):
            return
        if mInfo.Info.is_instance(value, BorderLine2Comp):
            self.__component.TopBorder2 = value.component
        else:
            self.__component.TopBorder2 = cast("BorderLine2", value)
        self.__props[key] = BorderLine2Comp(self.__component.TopBorder2, key, self.__event_provider)

    @property
    def user_defined_attributes(self) -> NameContainerComp | None:
        """
        Gets/Sets - stores additional attributes.

        This property is used i.e. by the XML filters to load and restore unknown attributes.
        """
        with contextlib.suppress(AttributeError):
            comp = self.__component.UserDefinedAttributes
            return None if comp is None else NameContainerComp(component=comp)
        return None

    @user_defined_attributes.setter
    def user_defined_attributes(self, value: NameContainerComp | XNameContainer) -> None:
        if hasattr(self.__component, "UserDefinedAttributes"):
            if mInfo.Info.is_instance(value, NameContainerComp):
                self.__component.UserDefinedAttributes = value.component
            else:
                self.__component.UserDefinedAttributes = value  # type: ignore

    @property
    def vert_justify(self) -> CellVertJustify2Enum:
        """
        Gets/Sets the vertical alignment of the cell contents.

        Hint:
            - ``CellVertJustify2Enum`` can be imported from ``ooo.dyn.table.cell_vert_justify2``
        """
        return CellVertJustify2Enum(self.__component.VertJustify)  # type: ignore

    @vert_justify.setter
    def vert_justify(self, value: int | CellVertJustify2Enum) -> None:
        val = CellVertJustify2Enum(value)
        self.__component.VertJustify = val.value

    # endregion Properties
