# region Imports
from __future__ import annotations
from typing import Any, Tuple, overload

from com.sun.star.graphic import XGraphic
from ooodev.utils import lo as mLo
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.meta.disabled_method import DisabledMethod

from ooodev.format.inner.common.props.img_para_area_props import ImgParaAreaProps
from ooodev.format.inner.direct.write.para.area.img import Img as ParaImg

# endregion Imports


class Img(ParaImg):
    """
    Class for table background image.

    .. versionadded:: 0.9.0
    """

    from_obj = DisabledMethod()
    """From object is not supported in this class."""

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextTable",)
        return self._supported_services_values

    def _on_multi_child_style_applying(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == "fill_image":
            event_args.cancel = True
        return super()._on_multi_child_style_applying(source, event_args)

    # region apply()
    @overload
    def apply(self, obj: object, **kwargs) -> None:
        ...

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        self._clear()
        self._set(self._props.back_color, -1)
        self._set(self._props.transparent, True)
        # will not work if not first converted to XGraphic
        graphic = mLo.Lo.qi(XGraphic, self.prop_inner.prop_bitmap, True)
        self._set(self._props.back_graphic, graphic)
        loc = self._get_graphic_loc(position=None, mode=self.prop_inner.prop_mode)
        self._set(self._props.graphic_loc, loc)
        super().apply(obj, **kwargs)

    # endregion apply()

    @property
    def _props(self) -> ImgParaAreaProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImgParaAreaProps(
                back_color="BackColor",
                back_graphic="BackGraphic",
                graphic_loc="BackGraphicLocation",
                transparent="BackTransparent",
            )
        return self._props_internal_attributes
