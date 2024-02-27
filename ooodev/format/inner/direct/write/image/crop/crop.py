"""
Module for Image Crop.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, Type, cast, TypeVar, overload, TYPE_CHECKING
import math

from ooo.dyn.text.graphic_crop import GraphicCrop

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.image_crop_props import ImageCropProps
from ooodev.format.inner.direct.structs.crop_struct import CropStruct
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.utils import images_lo as mImg
from ooodev.utils import props as mProps
from ooodev.utils.data_type.size import Size
from ooodev.utils.data_type.size_mm import SizeMM

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from ooodev.proto.size_obj import SizeObj
else:
    SizeObj = Any
    UnitT = Any

_TImageCrop = TypeVar(name="_TImageCrop", bound="ImageCrop")


class CropOpt(CropStruct):
    # region Init
    def __init__(
        self,
        *,
        left: float | UnitT = 0,
        right: float | UnitT = 0,
        top: float | UnitT = 0,
        bottom: float | UnitT = 0,
        all: float | UnitT | None = None,
        keep_scale: bool = True,
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitT, optional): Specifies left crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            right (float, UnitT, optional): Specifies right crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            top (float, UnitT, optional): Specifies top crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            bottom (float, UnitT, optional): Specifies bottom crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            keep_scale (bool, options): If ``True`` then crop is
            all (float, UnitT, optional): Specifies ``left``, ``right``, ``top``, and ``bottom`` in ``mm`` units or :ref:`proto_unit_obj`. If set all other parameters are ignored.
            keep_scale (bool, optional): If ``True`` crop is applied keeping image scale; Otherwise crop is applied keeping image size. Defaults to ``True``.
        """
        super().__init__(left=left, right=right, top=top, bottom=bottom, all=all)
        self._keep_scale = keep_scale

    # endregion Init

    # region Overrides
    def copy(self, **kwargs) -> CropOpt:
        cp = super().copy(**kwargs)
        cp._keep_scale = self._keep_scale
        return cp

    # endregion Overrides

    # region dunder methods
    def __eq__(self, oth: object) -> bool:
        result = super().__eq__(oth)
        if result == False:
            return False
        if isinstance(oth, CropOpt):
            return self.prop_keep_scale == oth.prop_keep_scale
        return False

    # endregion dunder methods

    # region Methods
    def can_crop(self) -> bool:
        """
        Gets if values are valid for crop.

        Returns:
            bool: ``True`` if options are valid; Otherwise, ``False``.
        """
        result = False
        for prop in self._props:
            pv = cast(int, self._get(prop))
            if pv != 0:
                result = True
                break
        return result

    # endregion Methods
    # region Overrides
    @property
    def prop_keep_scale(self) -> bool:
        """
        Gets/Sets keep scale.
        """
        return self._keep_scale

    @prop_keep_scale.setter
    def prop_keep_scale(self, value: bool):
        self._keep_scale = value

    # endregion Overrides


class ImageCrop(StyleMulti):
    """
    Crops and/or resizes an image.

    General Rules.

    **Crop Rules**

    Rules for ``CropOpt.keep_scale=True``.

    .. cssclass:: ul-list

        - If scale is passed in then image size is to be calculated from that scale using original image values, factoring in crop values.
        - If scale is not passed in image size is calculated from 100% using original image values, factoring in crop values.
        - If Image size is passed in then it is ignored.

    Rules for ``CropOpt.keep_scale=False``.

    .. cssclass:: ul-list

        - If image size is passed in then it is used to set image size, factoring in crop values.
        - If image size is not passed in then then the original image size is used to set image size, factoring in crop values.

    **No Crop Rules**

    Rules when crop is not passed to constructor.

    .. cssclass:: ul-list

        - If image size is present it is used to set the image size. In this case scale is ignored. Any existing Crop values are ignored.
        - If scale is present but no image size is present then a new size is derived from original size using scale.
        - If both image size and scale size are present then scale is ignored.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        crop: CropOpt | None = None,
        img_size: SizeMM | None = None,
        img_scale: SizeObj | None = None,
    ) -> None:
        """
        Constructor

        Args:
            crop: (CropOpt, optional): Specifies crop values.
            img_size (SizeMM, optional): Specifies image size.
            image_scale (SizeObj, optional): Specifies image scale.
        """
        super().__init__()

        self._img_size = img_size
        self._img_scale = None if img_scale is None else Size.from_size(img_scale)
        if crop is not None:
            co = crop.copy(_cattribs=self._get_struct_cattrib())
            self._set_style("crop_struct", co, *co.get_attrs())

    # endregion Init

    # region internal methods

    def _get_struct_cattrib(self) -> dict:
        return {
            "_property_name": self._props.crop_struct,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    def _rule_crop_keep_scale_scale(self, orig_size: Size) -> Size:
        # sourcery skip: class-extract-method
        # orig_size in 1/100th mm
        # returns 1/100th mm
        # If scale is passed in then image size is to be calculated from that scale using original image values, factoring in crop values
        img_scale = self.prop_img_scale
        if img_scale is None:
            raise ValueError("Image Scale is required for rule crop-scale.")
        crop = self.prop_crop_opt
        if crop is None:
            raise ValueError("Crop Options are required for rule crop-scale.")

        img_width = orig_size.width
        img_height = orig_size.height

        # do calculations in 1/100th mm
        width = mImg.ImagesLo.calc_keep_scale_len(
            orig_len=img_width,
            start_crop=crop.prop_top.get_value_mm100(),
            end_crop=crop.prop_bottom.get_value_mm100(),
            scale=img_scale.width / 100,
        )
        height = mImg.ImagesLo.calc_keep_scale_len(
            orig_len=img_height,
            start_crop=crop.prop_left.get_value_mm100(),
            end_crop=crop.prop_right.get_value_mm100(),
            scale=img_scale.height / 100,
        )
        # returns 1/100th mm
        return Size(width=round(width), height=round(height))

    def _rule_crop_keep_scale_no_scale(self, orig_size: Size) -> Size:
        # orig_size in 1/100th mm
        # returns 1/100th mm
        # If scale is not passed in image size is calculated from 100% using original image values, factoring in crop values.
        img_scale = Size(100, 100)
        crop = self.prop_crop_opt
        if crop is None:
            raise ValueError("Crop Options are required for rule crop-scale.")

        img_width = orig_size.width
        img_height = orig_size.height

        # do calculations in 1/100th mm
        width = mImg.ImagesLo.calc_keep_scale_len(
            orig_len=img_width,
            start_crop=crop.prop_top.get_value_mm100(),
            end_crop=crop.prop_bottom.get_value_mm100(),
            scale=img_scale.width / 100,
        )
        height = mImg.ImagesLo.calc_keep_scale_len(
            orig_len=img_height,
            start_crop=crop.prop_left.get_value_mm100(),
            end_crop=crop.prop_right.get_value_mm100(),
            scale=img_scale.height / 100,
        )
        # returns 1/100th mm
        return Size(width=round(width), height=round(height))

    def _rule_no_crop_image(self) -> Size:
        # returns 1/100th mm
        # image size is present use it. In this case scale is ignored
        if self.prop_img_size is None:
            raise ValueError("Crop Image is required for rule no-crop-image.")
        return self.prop_img_size.get_size_mm100()

    def _rule_no_crop_scale_no_image(self, orig_size: Size) -> Size:
        # orig_size in 1/100th mm
        # returns 1/100th mm
        # Scale No images size calculate new size from original size using scale.
        if self.prop_img_scale is None:
            raise ValueError("Crop Image is required for rule no-crop-scale-no-image.")
        factor_width = self.prop_img_scale.width / 100
        factor_height = self.prop_img_scale.height / 100
        new_width = orig_size.width * factor_width
        new_height = orig_size.height * factor_height
        return Size(width=round(new_width), height=round(new_height))

    def _rule_crop_keep_image(self, orig_size: Size) -> Size:
        # orig_size in 1/100th mm
        # returns 1/100th mm
        # If image size is passed in then it is use that the image size
        # If image is not passed in then then the original image size is used
        if self.prop_img_size is None:
            return orig_size
        return self.prop_img_size.get_size_mm100()

    def _get_keep_scale_value(self, orig_size: Size) -> Size:
        # orig_size in 1/100th mm
        # If scale is passed in then image size is to be calculated from that scale using original image values, factoring in crop values
        # If scale is not passed in image size is calculated from 100% using original image values, scale factoring in crop values.
        # If Image size is passed in then it is ignored.

        if self.prop_img_scale is not None:
            return self._rule_crop_keep_scale_scale(orig_size)
        return self._rule_crop_keep_scale_no_scale(orig_size)

    # endregion internal methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextGraphicObject",)
        return self._supported_services_values

    def _on_modifying(self, source: Any, event_args: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event_args)

    # region Copy()
    @overload
    def copy(self: _TImageCrop) -> _TImageCrop: ...

    @overload
    def copy(self: _TImageCrop, **kwargs) -> _TImageCrop: ...

    def copy(self: _TImageCrop, **kwargs) -> _TImageCrop:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._img_size = self._img_size
        cp._img_scale = self._img_scale
        return cp

    # endregion Copy(

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies style of current instance.

        Args:
            obj (Any): UNO Object that styles are to be applied.
        """
        # sourcery skip: de-morgan, hoist-statement-from-if, merge-else-if-into-elif, use-named-expression
        if not self._is_valid_obj(obj):
            self._print_not_valid_srv(method_name="apply")
            return

        actual_size = cast(SizeObj, mProps.Props.get(obj, self._props.actual_size))
        orig_size = Size.from_size(actual_size)

        apply_clear = kwargs.pop("_apply_clear", True)
        if apply_clear:
            self._clear()
        crop = self.prop_crop_opt
        if crop is None:
            sz = None
            if not self.prop_img_size is None:
                sz = self._rule_no_crop_image()
            elif not self.prop_img_scale is None:
                sz = self._rule_no_crop_scale_no_image(orig_size)
        else:
            if crop.prop_keep_scale:
                sz = self._get_keep_scale_value(orig_size)
            else:
                sz = self._rule_crop_keep_image(orig_size)
        if not sz is None:
            self._set(self._props.height, sz.height)
            self._set(self._props.width, sz.width)
        super().apply(obj, **kwargs)

    # endregion apply()

    # endregion Overrides

    # region Static Methods
    # region reset_to_original_size()
    @overload
    @classmethod
    def reset_image_original_size(cls, obj: object) -> None: ...

    @overload
    @classmethod
    def reset_image_original_size(cls, obj: object, **kwargs) -> None: ...

    @classmethod
    def reset_image_original_size(cls, obj: object, **kwargs) -> None:
        """
        Resets the image to its original size. Resetting crop, scale and size.

        Args:
            obj (object): UNO Image Object

        Raises:
            NotSupportedError: If ``obj`` is not supported.
        """
        inst = cls(crop=CropOpt(all=0), img_scale=Size(100, 100), **kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        actual_size = cast(SizeObj, mProps.Props.get(obj, inst._props.actual_size))
        inst.prop_img_size = SizeMM.from_size_mm100(actual_size)
        inst.apply(obj)

    # endregion reset_to_original_size()

    # region from_obj_get_size()

    @overload
    @classmethod
    def get_image_original_size(cls, obj: object) -> Size: ...

    @overload
    @classmethod
    def get_image_original_size(cls, obj: object, **kwargs) -> Size: ...

    @classmethod
    def get_image_original_size(cls, obj: object, **kwargs) -> Size:
        """
        Gets size from object in ``1/100th mm`` units.

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Size: Size in ``1/100th mm`` units.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        actual_size = cast(SizeObj, mProps.Props.get(obj, inst._props.actual_size))
        return Size.from_size(actual_size)

    # endregion from_obj_get_size()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TImageCrop], obj: object) -> _TImageCrop: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TImageCrop], obj: object, **kwargs) -> _TImageCrop: ...

    @classmethod
    def from_obj(cls: Type[_TImageCrop], obj: object, **kwargs) -> _TImageCrop:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Crop: Instance that represents Image crop.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        actual_size = cast(SizeObj, mProps.Props.get(obj, inst._props.actual_size))
        size = cast(SizeObj, mProps.Props.get(obj, inst._props.size))
        inst.prop_img_size = SizeMM.from_size_mm100(size)

        crop_struct = cast(GraphicCrop, mProps.Props.get(obj, inst._props.crop_struct))
        cs = CropOpt.from_uno_struct(value=crop_struct, _cattribs=inst._get_struct_cattrib())

        # get scale
        scale_W = mImg.ImagesLo.calc_scale_crop(
            orig_len=actual_size.Width, new_len=size.Width, start_crop=crop_struct.Left, end_crop=crop_struct.Right
        )
        scale_h = mImg.ImagesLo.calc_scale_crop(
            orig_len=actual_size.Height, new_len=size.Height, start_crop=crop_struct.Top, end_crop=crop_struct.Bottom
        )
        inst.prop_img_scale = Size(math.ceil(scale_W), math.ceil(scale_h))

        sz_actual = Size.from_size(actual_size)
        sz_size = Size.from_size(size)
        cs.prop_keep_scale = sz_actual != sz_size
        inst._set_style("crop_struct", cs, *cs.get_attrs())
        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.IMAGE
        return self._format_kind_prop

    @property
    def prop_img_size(self) -> SizeMM | None:
        """
        Gets or Sets image size.
        """
        return self._img_size

    @prop_img_size.setter
    def prop_img_size(self, value: SizeMM | None):
        self._img_size = value

    @property
    def prop_img_scale(self) -> Size | None:
        """
        Gets or Sets image scale.
        """
        return self._img_scale

    @prop_img_scale.setter
    def prop_img_scale(self, value: SizeObj | None):
        if value is None:
            self._img_scale = None
        else:
            sz = Size.from_size(value)
            sz.height = max(sz.height, 1)
            sz.width = max(sz.width, 1)
            self._img_scale = sz

    @property
    def prop_crop_opt(self) -> CropOpt | None:
        """Gets or Sets Crop Struct instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(CropOpt, self._get_style_inst("crop_struct"))
        return self._direct_inner

    @prop_crop_opt.setter
    def prop_crop_opt(self, value: CropOpt | None) -> None:
        self._del_attribs("_direct_inner")
        if value is None:
            self._remove_style("crop_struct")
            return
        self._set_style("crop_struct", value, *value.get_attrs())

    @property
    def _props(self) -> ImageCropProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImageCropProps(
                crop_struct="GraphicCrop", width="Width", height="Height", size="Size", actual_size="ActualSize"
            )
        return self._props_internal_attributes

    # endregion Properties
