from __future__ import annotations
from typing import cast, overload
from typing import Any, Tuple, Type, TypeVar
from enum import Enum
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.image_flip_props import ImageFlipProps

_TFlip = TypeVar("_TFlip", bound="Flip")


class FlipKind(Enum):
    """Horizontal Flip Options"""

    NONE = 0
    """No flip"""
    ALL = 1
    """On all pages"""
    LEFT = 2
    """"On left pages"""
    RIGHT = 3
    """On right pages"""


class Flip(StyleBase):
    """
    Image Flip

    .. versionadded:: 0.9.0
    """

    def __init__(self, vertical: bool | None = None, horizontal: FlipKind | None = None) -> None:
        """
        Constructor

        Args:
            printable (bool, optional): Specifies if Frame can be printed. Default ``True``.
        """
        super().__init__()
        if vertical is not None:
            self.prop_vertical = vertical
        if horizontal is not None:
            self.prop_Horizontal = horizontal

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
    def from_obj(cls: Type[_TFlip], obj: object) -> _TFlip: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TFlip], obj: object, **kwargs) -> _TFlip: ...

    @classmethod
    def from_obj(cls: Type[_TFlip], obj: object, **kwargs) -> _TFlip:
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
    def prop_vertical(self) -> bool | None:
        """Gets/Sets Vertical flip option"""
        return self._get(self._props.vert)

    @prop_vertical.setter
    def prop_vertical(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.vert)
            return
        self._set(self._props.vert, value)

    @property
    def prop_Horizontal(self) -> FlipKind | None:
        """Gets/Sets Horizontal flip option"""
        even = cast(bool, self._get(self._props.hori_even))
        odd = cast(bool, self._get(self._props.hori_odd))
        if even is None or odd is None:
            return None

        if even:
            return FlipKind.ALL if odd else FlipKind.LEFT
        return FlipKind.RIGHT if odd else FlipKind.NONE

    @prop_Horizontal.setter
    def prop_Horizontal(self, value: FlipKind | None) -> None:
        if value is None:
            self._remove(self._props.hori_even)
            self._remove(self._props.hori_odd)
            return
        if value == FlipKind.ALL:
            self._set(self._props.hori_even, True)
            self._set(self._props.hori_odd, True)
        elif value == FlipKind.LEFT:
            self._set(self._props.hori_even, True)
            self._set(self._props.hori_odd, False)
        elif value == FlipKind.RIGHT:
            self._set(self._props.hori_even, False)
            self._set(self._props.hori_odd, True)
        else:
            self._set(self._props.hori_even, False)
            self._set(self._props.hori_odd, False)

    @property
    def _props(self) -> ImageFlipProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = ImageFlipProps(
                vert="VertMirrored", hori_even="HoriMirroredOnEvenPages", hori_odd="HoriMirroredOnOddPages"
            )
        return self._props_internal_attributes

    # endregion Properties
