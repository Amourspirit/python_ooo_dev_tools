from __future__ import annotations
from typing import Dict, Set, TYPE_CHECKING

import uno

from ooodev.io.log import logging as logger
from ooodev.calc.cell.custom_prop_clean import CustomPropClean
from ooodev.utils.data_type.cell_obj import CellObj

if TYPE_CHECKING:
    from ooodev.calc.calc_sheet import CalcSheet


class SheetCellCustomProperties(CustomPropClean):
    def __init__(self, sheet: CalcSheet):
        CustomPropClean.__init__(self, sheet)

    def get_cell_properties(self, *filter: str) -> Dict[CellObj, Set[str]]:
        """
        Get Custom Property Cells

        The keys are sorted in the cell order of across and then down. The values are the custom property names.

        Args:
            filter (Tuple[str], optional): Custom Property Names to filter by. If omitted all custom properties are returned.
                Otherwise only the cell with custom properties in the list are returned.

        Returns:
            Dict[mCellObj.CellObj, Set[str]]: Cell and Custom Property Names
        """
        # Get all cell that have custom properties and their property names

        self.clean()
        all_shapes = self._get_shapes_artifact_dict()
        prop_shapes = all_shapes["prop_shapes"]
        result = {}
        if filter:
            filter_set = set(filter)
        else:
            filter_set = None
        for _, shapes in prop_shapes.items():
            shape = shapes[0]
            anchor = shape.Anchor  # type: ignore
            # self.clean() would have removed the shape if the anchor is None or if cell has been deleted
            assert anchor is not None

            cell_obj = CellObj.from_cell(anchor.CellAddress)

            current_set = set()
            cell = self._sheet[cell_obj]
            if cell.has_custom_properties():
                props = cell.get_custom_properties()
                for key in props.keys():
                    current_set.add(key)
            # check of at least one of the properties is in the filter
            if filter_set is None or current_set.intersection(filter_set):
                result[cell_obj] = current_set

        return dict(sorted(result.items()))

    def remove_all_custom_properties(self) -> None:
        """
        Remove all custom properties from all cells.

        Warning:
            This is a destructive operation. It will remove all custom properties from all cells for the entire sheet.

            Only use this if you are sure you want to remove all custom properties and you know what you are doing.
            It is possible other extensions or macros may rely on custom properties.
        """
        self.clean()
        all_shapes = self._get_shapes_artifact_dict()
        prop_shapes = all_shapes["prop_shapes"]
        draw_page = self._sheet.draw_page
        for _, shapes in prop_shapes.items():
            for shape in shapes:
                logger.debug(f"Removing shape {shape.Name}")  # type: ignore
                draw_page.remove(shape)
        if draw_page.forms.has_by_name(self.form_name):
            logger.debug(f"Removing form {self.form_name}")
            draw_page.forms.remove_by_name(self.form_name)
