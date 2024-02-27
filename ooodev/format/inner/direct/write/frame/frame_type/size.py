"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, cast, overload
from ooo.dyn.text.size_type import SizeTypeEnum

from ooodev.utils import props as mProps
from ooodev.format.inner.common.props.frame_type_size_props import FrameTypeSizeProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.direct.write.image.image_type.size import Size as ImageSize
from ooodev.format.inner.direct.write.image.image_type.size import RelativeSize
from ooodev.format.inner.direct.write.image.image_type.size import AbsoluteSize

_TSize = TypeVar(name="_TSize", bound="Size")


class Size(ImageSize):
    """
    Frame Type Size

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        width: RelativeSize | AbsoluteSize | None = None,
        height: RelativeSize | AbsoluteSize | None = None,
        auto_width: bool = False,
        auto_height: bool = False,
    ) -> None:
        """
        Constructor

        Args:
            width (RelativeSize, AbsoluteSize, optional): width value.
            height (RelativeSize, AbsoluteSize, optional): height value.
            auto_width (bool, optional): Auto Size Width. Default ``False``.
            auto_height (bool, optional): Auto Size Height. Default ``False``.
        """
        # size width as a percent is max value of 254
        # Width
        #   Relative to entire page ((PageWidth - (LeftMargin - RightMargin)) x percent)

        # RelativeWidthRelation is RelativeKind value
        # RelativeHeighRelation is RelativeKind value
        # When AbsoluteSize RelativeWidth=0, Size.Width=(width 1/100 mm) Width=(width 1/100 mm)
        # When RelativeSize = RelativeWidth=RelativeSize.size, Size.Width and Width become a calculated value base upon RelativeSize.Kind
        super().__init__(width=width, height=height)
        self._auto_width = auto_width
        self._auto_height = auto_height

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.Style",
                "com.sun.star.text.TextFrame",
            )
        return self._supported_services_values

    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies style of current instance.

        Args:
            obj (object): UNO Object that styles are to be applied.
        """
        if kwargs.pop("_apply_clear", True):
            self._clear()
        if self.prop_width:
            if self._auto_width:
                self._set(self._props.width_type, SizeTypeEnum.MIN.value)
            else:
                self._set(self._props.width_type, SizeTypeEnum.FIX.value)

        if self.prop_height:
            if self._auto_height:
                self._set(self._props.size_type, SizeTypeEnum.MIN.value)
            else:
                self._set(self._props.size_type, SizeTypeEnum.FIX.value)

        super().apply(obj=obj, _apply_clear=False, **kwargs)

    # region copy()
    @overload
    def copy(self: _TSize) -> _TSize: ...

    @overload
    def copy(self: _TSize, **kwargs) -> _TSize: ...

    def copy(self: _TSize, **kwargs) -> _TSize:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._auto_width = self._auto_width
        cp._auto_height = self._auto_height
        return cp

    # endregion copy()

    # endregion Overrides

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSize], obj: Any) -> _TSize: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSize], obj: Any, **kwargs) -> _TSize: ...

    @classmethod
    def from_obj(cls: Type[_TSize], obj: Any, **kwargs) -> _TSize:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Returns:
            Size: Instance that represents Frame size.
        """
        inst = cast(_TSize, super(cls, cls)).from_obj(obj, **kwargs)
        # https://tinyurl.com/2mdozjx2

        auto_width = SizeTypeEnum(int(mProps.Props.get(obj, inst._props.width_type, SizeTypeEnum.FIX.value)))
        inst._auto_width = auto_width == SizeTypeEnum.MIN
        auto_height = SizeTypeEnum(int(mProps.Props.get(obj, inst._props.size_type, SizeTypeEnum.FIX.value)))
        inst._auto_height = auto_height == SizeTypeEnum.MIN
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
            self._format_kind_prop = FormatKind.FRAME
        return self._format_kind_prop

    @property
    def prop_width(self) -> RelativeSize | AbsoluteSize | None:
        """
        Gets/Sets width.
        """
        return self._width

    @prop_width.setter
    def prop_width(self, value: RelativeSize | AbsoluteSize | None):
        self._width = value

    @property
    def prop_height(self) -> RelativeSize | AbsoluteSize | None:
        """
        Gets/Sets height.
        """
        return self._height

    @property
    def prop_auto_width(self) -> bool:
        """
        Gets/Sets auto width.
        """
        return self._auto_width

    @prop_auto_width.setter
    def prop_auto_width(self, value: bool):
        self._auto_width = value

    @property
    def prop_auto_height(self) -> bool:
        """
        Gets/Sets auto height.
        """
        return self._auto_height

    @prop_auto_height.setter
    def prop_auto_height(self, value: bool):
        self._auto_height = value

    @prop_height.setter
    def prop_height(self, value: RelativeSize | AbsoluteSize | None):
        self._height = value

    @property
    def _props(self) -> FrameTypeSizeProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FrameTypeSizeProps(
                width="Width",
                height="Height",
                rel_width="RelativeWidth",
                rel_height="RelativeHeight",
                rel_width_relation="RelativeWidthRelation",
                rel_height_relation="RelativeHeightRelation",
                size_type="SizeType",
                width_type="WidthType",
            )
        return self._props_internal_attributes

    # endregion Properties
