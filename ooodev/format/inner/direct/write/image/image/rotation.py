from __future__ import annotations
from typing import cast, overload
from typing import Any, Tuple, Type, TypeVar

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.image_rotation_props import ImageRotationProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.units.angle import Angle
from ooodev.utils import props as mProps

_TRotation = TypeVar(name="_TRotation", bound="Rotation")


class Rotation(StyleBase):
    """
    Image Rotation

    .. versionadded:: 0.9.0
    """

    def __init__(self, rotation: int | Angle = 0) -> None:
        """
        Constructor

        Args:
            rotation (int, Angle, optional): Specifies if the image rotation. Default ``0``.
        """
        super().__init__()
        self.prop_rotation = rotation

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextGraphicObject",)
        return self._supported_services_values

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TRotation], obj: object) -> _TRotation: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TRotation], obj: object, **kwargs) -> _TRotation: ...

    @classmethod
    def from_obj(cls: Type[_TRotation], obj: object, **kwargs) -> _TRotation:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Properties: Instance that represents Image Flip Options.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        for prop_name in inst._props:
            if prop_name:
                inst._set(prop_name, mProps.Props.get(obj, prop_name))
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
    def prop_rotation(self) -> Angle:
        """
        Gets/Sets Rotation angle of the shape in degrees.

        Property can be set by passing int or Angle.
        """
        # in 1/10 degree units
        pv = cast(int, self._get(self._props.rotation))
        if pv == 0:
            return Angle(0)
        return Angle(round(pv / 10))

    @prop_rotation.setter
    def prop_rotation(self, value: int | Angle) -> None:
        # in 1/10 degree units
        val = Angle(int(value)).value * 10
        self._set(self._props.rotation, val)

    @property
    def _props(self) -> ImageRotationProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImageRotationProps(rotation="GraphicRotation")
        return self._props_internal_attributes

    # endregion Properties
