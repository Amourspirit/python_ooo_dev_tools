"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
import dataclasses
from typing import Any, Tuple, Type, TypeVar, overload
from enum import Enum
import math
import uno
from ooo.dyn.text.size_type import SizeTypeEnum

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.unit_convert import UnitConvert
from .....utils.validation import check
from ....direct.common.abstract.abstract_document import AbstractDocument
from ....direct.common.props.frame_type_size_props import FrameTypeSizeProps
from ....kind.format_kind import FormatKind


_TSize = TypeVar(name="_TSize", bound="Size")


class RelativeKind(Enum):
    """Relative Kind"""

    PAGE = 7
    """Enitre Page"""
    PARAGRAPH = 0
    """Paragraph area"""

    def __int__(self) -> int:
        return self.value


@dataclasses.dataclass(frozen=True)
class RelativeSize:
    """Relative size"""

    size: int
    """Percentage of Page Or Paragraph from 1 to 254"""
    kind: RelativeKind
    """Relative Kind"""

    def __post_init__(self) -> None:
        check(
            self.size >= 1 and self.size <= 254,
            f"{self}",
            f"Value of {self.size} is out of range. Value must be from 2 to 254.",
        )

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, RelativeSize):
            return self.size == oth.size and self.kind.value == oth.kind.value
        return NotImplemented


@dataclasses.dataclass(frozen=True)
class AbsoluteSize:
    """Absolute size"""

    size: float
    """Size in ``mm`` units"""

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, AbsoluteSize):
            return math.isclose(self.size, oth.size, abs_tol=0.02)
        if isinstance(oth, float):
            return math.isclose(self.size, oth, abs_tol=0.02)
        return NotImplemented


class Size(AbstractDocument):
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
        #   Realitive to eniter page ((PageWidth - (LeftMargin - RightMargin)) x percent)

        # RelativeWidthRelation is RelativeKind value
        # RelativeHeighRelation is RelativeKind value
        # When AbsoluteSize RealitiveWidth=0, Size.Width=(width 1/100 mm) Width=(width 1/100 mm)
        # When RelativeSize = RealitiveWidth=RelativeSize.size, Size.Width and Width become a caclulated value base upon RelativeSize.Kind
        super().__init__()
        self._width = width
        self._height = height
        self._auto_width = auto_width
        self._auto_height = auto_height

    # region Internal Methods

    def _get_rel_width(self, size: RelativeSize) -> int:
        if size.kind == RelativeKind.PAGE:
            doc_size = self.get_page_size()
        else:
            doc_size = self.get_page_text_size()
        percent = size.size / 100
        return round(doc_size.Width * percent)

    def _get_rel_height(self, size: RelativeSize) -> int:
        if size.kind == RelativeKind.PAGE:
            doc_size = self.get_page_size()
        else:
            doc_size = self.get_page_text_size()
        percent = size.size / 100
        return round(doc_size.Height * percent)

    # endregion Internal Methods

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.Style",)
        return self._supported_services_values

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @overload
    def apply(self, obj: object) -> None:
        pass

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies style of current instance and all other internal style instances.

        Args:
            obj (object): UNO Oject that styles are to be applied.
        """
        self._clear()
        if self._width is None and self._height is None:
            return
        if self._width:
            if isinstance(self._width, AbsoluteSize):
                self._set(self._props.width, UnitConvert.convert_mm_mm100(self._width.size))
                self._set(self._props.rel_width, 0)

            else:
                if self._width.size > 1:
                    # relative width has a minimum of 2
                    rel_width = self._get_rel_width(self._width)
                    self._set(self._props.rel_width, self._width.size)
                else:
                    rel_width = self._get_rel_width(RelativeSize(2, kind=self._width.kind))
                    self._set(self._props.rel_width, 2)

                self._set(self._props.width, rel_width)
                self._set(self._props.rel_width_relation, self._width.kind.value)

            if self._auto_width:
                self._set(self._props.width_type, SizeTypeEnum.MIN.value)
            else:
                self._set(self._props.width_type, SizeTypeEnum.FIX.value)

        if self._height:
            if isinstance(self._height, AbsoluteSize):
                self._set(self._props.height, UnitConvert.convert_mm_mm100(self._height.size))
                self._set(self._props.rel_height, 0)

            else:
                rel_height = self._get_rel_height(self._height)
                self._set(self._props.height, rel_height)
                self._set(self._props.rel_height, self._height.size)
                self._set(self._props.rel_height_relation, self._height.kind.value)

            if self._auto_height:
                self._set(self._props.size_type, SizeTypeEnum.MIN.value)
            else:
                self._set(self._props.size_type, SizeTypeEnum.FIX.value)

        return super().apply(obj, **kwargs)

    def copy(self: _TSize) -> _TSize:
        cp = super().copy()
        if self._width is None:
            cp._width = None
        else:
            cp._width = dataclasses.replace(self._width)
        if self._height is None:
            cp._height = None
        else:
            cp._height = dataclasses.replace(self._height)
        cp._auto_width = self._auto_width
        cp._auto_height = self._auto_height
        return cp

    # endregion Overrides

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSize], obj: object) -> _TSize:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSize], obj: object, **kwargs) -> _TSize:
        ...

    @classmethod
    def from_obj(cls: Type[_TSize], obj: object, **kwargs) -> _TSize:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Returns:
            Size: Instance that represents Frame size.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        width = int(mProps.Props.get(obj, inst._props.width))
        rel_width = int(mProps.Props.get(obj, inst._props.rel_width, RelativeKind.PARAGRAPH.value))
        rel_width_relation = RelativeKind(
            int(mProps.Props.get(obj, inst._props.rel_width_relation, RelativeKind.PARAGRAPH.value))
        )
        is_abs_width = rel_width == 0
        if is_abs_width:
            inst._width = AbsoluteSize(round(UnitConvert.convert_mm100_mm(width), 2))
        else:
            inst._width = RelativeSize(rel_width, rel_width_relation)

        height = int(mProps.Props.get(obj, inst._props.height))
        rel_height = int(mProps.Props.get(obj, inst._props.rel_height, RelativeKind.PARAGRAPH.value))
        rel_height_relation = RelativeKind(
            int(mProps.Props.get(obj, inst._props.rel_height_relation, RelativeKind.PARAGRAPH.value))
        )
        is_abs_height = rel_height == 0
        if is_abs_height:
            inst._height = AbsoluteSize(round(UnitConvert.convert_mm100_mm(height), 2))
        else:
            inst._height = RelativeSize(rel_height, rel_height_relation)

        auto_width = SizeTypeEnum(int(mProps.Props.get(obj, inst._props.width_type, SizeTypeEnum.FIX.value)))
        if auto_width == SizeTypeEnum.MIN:
            inst._auto_width = True
        else:
            inst._auto_width = False

        auto_height = SizeTypeEnum(int(mProps.Props.get(obj, inst._props.size_type, SizeTypeEnum.FIX.value)))
        if auto_height == SizeTypeEnum.MIN:
            inst._auto_height = True
        else:
            inst._auto_height = False

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
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
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
