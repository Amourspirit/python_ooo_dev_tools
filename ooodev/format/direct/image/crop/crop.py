"""
Module for Image Crop.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, Type, cast, TypeVar, overload

import uno
from ooo.dyn.text.graphic_crop import GraphicCrop

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....proto.unit_obj import UnitObj
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleMulti
from ...structs.crop_struct import CropStruct
from ...common.props.image_crop_props import ImageCropProps


_TCrop = TypeVar(name="_TCrop", bound="Crop")


class Crop(StyleMulti):
    """
    Image Crop

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        left: float | UnitObj = 0.0,
        right: float | UnitObj = 0.0,
        top: float | UnitObj = 0.0,
        bottom: float | UnitObj = 0.0,
        all: float | UnitObj = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitObj, optional): Specifies left crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            right (float, UnitObj, optional): Specifies right crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            top (float, UnitObj, optional): Specifies top crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            bottom (float, UnitObj, optional): Specifies bottom crop in ``mm`` units or :ref:`proto_unit_obj`. Default ``0.0``.
            all (float, UnitObj, optional): Specifies ``left``, ``right``, ``top``, and ``bottom`` in ``mm`` units or :ref:`proto_unit_obj`. If set all other paramaters are ignored.
        """
        cs = CropStruct(
            left=left,
            right=right,
            top=top,
            bottom=bottom,
            all=all,
            _cattribs=self._get_struct_cattrib(),
        )
        super().__init__()
        self._set_style("crop_struct", cs, *cs.get_attrs())

    # endregion Init

    # region internal methods

    def _get_struct_cattrib(self) -> dict:
        return {
            "_property_name": self._props.crop_struct,
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
        }

    # endregion internal methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextGraphicObject",)
        return self._supported_services_values

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    # endregion Overrides

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TCrop], obj: object) -> _TCrop:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TCrop], obj: object, **kwargs) -> _TCrop:
        ...

    @classmethod
    def from_obj(cls: Type[_TCrop], obj: object, **kwargs) -> _TCrop:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Crop: Instance that represents Image crop.
        """
        # this nu is only used to get Property Name
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        crop_struct = cast(GraphicCrop, mProps.Props.get(obj, inst._props.crop_struct))
        cs = CropStruct.from_uno_struct(crop_struct, _cattribs=inst._get_struct_cattrib())

        inst._set_style("crop_struct", cs, *cs.get_attrs())
        return inst

    # endregion from_obj()

    # region from_struct()
    @classmethod
    def from_struct(cls: Type[_TCrop], struct: CropStruct, **kwargs) -> _TCrop:
        """
        Gets instance from ``CropStruct`` instance

        Args:
            struct (CropStruct): Gradient Struct instance.
            name (str, optional): Name of Gradient.

        Returns:
            Gradient:
        """
        inst = cls(**kwargs)
        crop_struct = struct.get_uno_struct()
        cs = CropStruct.from_uno_struct(crop_struct, _cattribs=inst._get_struct_cattrib())
        inst._set_style("crop_struct", cs, *cs.get_attrs())
        return inst

    # endregion from_struct()

    # endregion Static Methods

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.IMAGE
        return self._format_kind_prop

    @property
    def prop_inner(self) -> CropStruct:
        """Gets Crop Struct instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(CropStruct, self._get_style_inst("crop_struct"))
        return self._direct_inner

    @property
    def _props(self) -> ImageCropProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImageCropProps(crop_struct="GraphicCrop")
        return self._props_internal_attributes
