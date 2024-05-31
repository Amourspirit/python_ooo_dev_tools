from __future__ import annotations
from typing import Any, Dict, List, cast, Set, TYPE_CHECKING
import contextlib

import uno
from com.sun.star.drawing import XControlShape
from com.sun.star.uno import RuntimeException

from ooodev.io.log import logging as logger
from ooodev.calc.cell.custom_prop_base import CustomPropBase

if TYPE_CHECKING:
    from ooodev.calc.calc_sheet import CalcSheet


class CustomPropClean(CustomPropBase):
    def __init__(self, sheet: CalcSheet):
        CustomPropBase.__init__(self, sheet)
        self._sheet = sheet

    def _get_shapes_artifact_dict(self) -> Dict[str, Dict[str, List[XControlShape]]]:

        comp = self.draw_page.component
        shapes = {
            "prop_shapes": {},
            "artifacts": {},
        }
        # find all shapes on the draw page that start with prefix and end with suffix
        for shape in comp:  # type: ignore
            if not shape.supportsService("com.sun.star.drawing.ControlShape"):
                continue
            name = cast(str, shape.Name)
            if name.startswith(self._shape_prefix):
                if name.endswith(self._shape_suffix):
                    if name in shapes["prop_shapes"]:
                        shapes["prop_shapes"][name].append(shape)
                    else:
                        shapes["prop_shapes"][name] = [shape]
                else:
                    if name in shapes["artifacts"]:
                        shapes["artifacts"][name].append(shape)
                    else:
                        shapes["artifacts"][name] = [shape]
        return shapes

    def _is_cell_deleted(self, cell: Any) -> bool:
        try:
            _ = cell.AbsoluteName
        except RuntimeException:
            return True
        return False

    def _has_hidden_control(self, shape: XControlShape) -> bool:
        ctl = self._get_hidden_control_simple(self._get_hidden_control_name_from_shape(shape))
        return ctl is not None

    def _clean_duplicate_shapes(self, shapes: List[XControlShape]) -> List[XControlShape]:
        # remove any duplicates named shapes
        if len(shapes) < 2:
            return shapes
        shapes.sort(key=lambda x: x.ZOrder)  # type: ignore

        for shape in shapes[1:]:
            with contextlib.suppress(Exception):
                logger.debug(f"Removing duplicate shape {shape.Name}")  # type: ignore
                self.draw_page.remove(shape)
        shapes = shapes[:1]
        return shapes

    def _get_all_hidden_controls(self) -> Set[str]:
        # get all hidden controls on the draw page
        hidden_controls = set()
        form = self._get_form()
        for control in form:  # type: ignore
            if control.supportsService("com.sun.star.form.component.HiddenControl"):
                hidden_controls.add(control.Name)
        return hidden_controls

    def _remove_hidden_controls(self, hidden_controls: Set[str]) -> None:
        # remove any hidden controls that are not associated with a shape.
        form = self._get_form()
        for control_name in hidden_controls:
            if form.hasByName(control_name):
                logger.debug(f"Removing hidden control {control_name}")  # type: ignore
                form.removeByName(control_name)

    def clean(self):
        # clean up any artifacts such a name like '_cprop_idofhsvtcky1hgom_id 1'
        # remove any duplicates named shapes
        # remove any shapes that have a cell anchor that is deleted.

        all_shapes = self._get_shapes_artifact_dict()
        for shapes in all_shapes["artifacts"].values():
            for shape in shapes:
                with contextlib.suppress(Exception):
                    logger.debug(f"Removing artifact shape {shape.Name}")  # type: ignore
                    self.draw_page.remove(shape)
        prop_shapes = all_shapes["prop_shapes"]
        # each shape should have a corresponding anchor that is a cell.

        for shape_name, shapes in prop_shapes.items():
            single_shapes = self._clean_duplicate_shapes(shapes)
            if not single_shapes:
                continue
            # should only have one shape at this point
            if len(single_shapes) > 1:
                logger.error(
                    f"{self.__class__.__name__} - More than one shape found for {shape_name}. Unable to clean."
                )
                continue
            # make sure the single_shape cell has not been deleted.
            shape = single_shapes[0]
            anchor = shape.Anchor  # type: ignore
            if anchor is None:
                with contextlib.suppress(Exception):
                    logger.debug(f"Removing shape {shape.Name} because anchor is None")  # type: ignore
                    self.draw_page.remove(shape)
                continue

            if self._is_cell_deleted(anchor):
                with contextlib.suppress(Exception):
                    logger.debug(f"Removing shape {shape.Name} because cell is deleted")  # type: ignore
                    self.draw_page.remove(shape)

        current_hidden_controls = self._get_all_hidden_controls()
        keep_hidden_controls = set()
        all_shapes = self._get_shapes_artifact_dict()
        prop_shapes = all_shapes["prop_shapes"]
        for _, shape in prop_shapes.items():
            keep_hidden_controls.add(self._get_hidden_control_name_from_shape(shape[0]))

        removed_ids = current_hidden_controls - keep_hidden_controls

        if removed_ids:
            self._remove_hidden_controls(removed_ids)
