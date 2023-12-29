from __future__ import annotations
from typing import Any, TypeVar, List, Protocol, overload, TYPE_CHECKING

from ooodev.adapter.container.name_container_comp import NameContainerComp
from ooodev.draw import draw_pages as mDrawPages
from ooodev.draw import master_draw_page as mMasterDrawPage
from ooodev.draw.shapes import draw_shape as mDrawShape
from ooodev.proto.component_proto import ComponentT
from ooodev.utils.kind.drawing_layer_kind import DrawingLayerKind
from ooodev.utils.kind.shape_comb_kind import ShapeCombKind
from ooodev.utils.kind.zoom_kind import ZoomKind

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.drawing import XDrawPages
    from com.sun.star.drawing import XDrawView
    from com.sun.star.drawing import XLayer
    from com.sun.star.drawing import XLayerManager
    from com.sun.star.drawing import XShapes
    from com.sun.star.frame import XController
    from ooodev.utils.data_type.size import Size

_T = TypeVar("_T", bound="ComponentT")


class DrawDocT(Protocol[_T]):
    component: Any

    # def add_slide(self) -> mDrawPage.DrawPage[_T]:
    def add_slide(self) -> Any:
        ...

    def add_layer(self, lm: XLayerManager, layer_name: str) -> XLayer:
        ...

    def build_play_list(self, custom_name: str, *slide_idxs: int) -> NameContainerComp:
        ...

    def close_doc(self, deliver_ownership: bool = ...) -> bool:
        ...

    def combine_shape(self, shapes: XShapes, combine_op: ShapeCombKind) -> mDrawShape.DrawShape[_T]:
        ...

    def delete_slide(self, idx: int) -> bool:
        ...

    def duplicate(self, idx: int) -> Any:
        ...

    def find_master_page(self, style: str) -> mMasterDrawPage.MasterDrawPage[_T]:
        ...

    def find_slide_idx_by_name(self, name: str) -> int:
        ...

    def get_controller(self) -> XController:
        ...

    def get_handout_master_page(self) -> Any:
        ...

    def get_layer(self, layer_name: DrawingLayerKind | str) -> XLayer:
        ...

    def get_layer_manager(self) -> XLayerManager:
        ...

    def get_master_page(self, idx: int) -> mMasterDrawPage.MasterDrawPage[_T]:
        ...

    def get_master_page_count(self) -> int:
        ...

    def get_ordered_shapes(self) -> List[mDrawShape.DrawShape[_T]]:
        ...

    def get_play_list(self) -> NameContainerComp:
        ...

    def get_shapes(self) -> List[mDrawShape.DrawShape[_T]]:
        ...

    def get_shapes_text(self) -> str:
        ...

    # region get_slide()
    @overload
    def get_slide(self) -> Any:
        ...

    @overload
    def get_slide(self, *, idx: int) -> Any:
        ...

    @overload
    def get_slide(self, *, slides: XDrawPages) -> Any:
        ...

    @overload
    def get_slide(self, *, slides: XDrawPages, idx: int) -> Any:
        ...

    # endregion get_slide()

    def get_slide_number(self, xdraw_view: XDrawView) -> int:
        ...

    def get_slide_size(self) -> Size:
        ...

    def get_slides(self) -> mDrawPages.DrawPages:
        ...

    def get_slides_count(self) -> int:
        ...

    def get_slides_list(self) -> list:
        ...

    def get_viewed_page(self) -> Any:
        ...

    def goto_page(self, page: XDrawPage) -> None:
        ...

    def insert_master_page(self, idx: int) -> mMasterDrawPage.MasterDrawPage[_T]:
        ...

    def insert_slide(self, idx: int) -> Any:
        ...

    def set_visible(self, visible: bool = True) -> None:
        ...

    def zoom(self, type: ZoomKind = ...) -> None:
        ...

    def zoom_value(self, value: int = ...) -> None:
        ...
