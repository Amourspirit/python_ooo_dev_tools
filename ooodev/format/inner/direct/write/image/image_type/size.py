"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
import dataclasses
from typing import Any, Tuple, Type, TypeVar, overload, TYPE_CHECKING
from enum import Enum
import math

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.abstract.abstract_document import AbstractDocument
from ooodev.format.inner.common.props.frame_type_size_props import FrameTypeSizeProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.loader import lo as mLo
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps
from ooodev.utils.validation import check

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

_TSize = TypeVar(name="_TSize", bound="Size")


class RelativeKind(Enum):
    """Relative Kind"""

    PAGE = 7
    """Entire Page"""
    PARAGRAPH = 0
    """Paragraph area"""

    def __int__(self) -> int:
        return self.value


@dataclasses.dataclass(frozen=True)
class RelativeSize:
    """Relative size"""

    size: int
    """Percentage of Page Or Paragraph from ``1`` to ``254``"""
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


class AbsoluteSize:
    """Absolute size"""

    def __init__(self, value: float | UnitT) -> None:
        """
        Constructor

        Args:
            value (float, UnitT): Size value in ``mm`` units or :ref:`proto_unit_obj`.
        """
        self.size = value

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, AbsoluteSize):
            return math.isclose(self.size.value, oth.size.value, abs_tol=0.02)
        if isinstance(oth, float):
            return math.isclose(self.size.value, oth, abs_tol=0.02)
        return NotImplemented

    def copy(self) -> AbsoluteSize:
        """Creates a copy of this object"""
        return AbsoluteSize(self.size)

    @property
    def size(self) -> UnitMM:
        """Gets/Sets Size value in ``mm`` units"""
        return UnitMM(self._size)

    @size.setter
    def size(self, value: float | UnitT):
        try:
            self._size = round(value.get_value_mm(), 2)  # type: ignore
        except AttributeError:
            self._size = float(value)  # type: ignore


class Size(AbstractDocument):
    """
    Image Type Size

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        width: RelativeSize | AbsoluteSize | None = None,
        height: RelativeSize | AbsoluteSize | None = None,
    ) -> None:
        """
        Constructor

        Args:
            width (RelativeSize, AbsoluteSize, optional): width value.
            height (RelativeSize, AbsoluteSize, optional): height value.
        """
        # size width as a percent is max value of 254
        # Width
        #   Relative to entire page ((PageWidth - (LeftMargin - RightMargin)) x percent)

        # RelativeWidthRelation is RelativeKind value
        # RelativeHeighRelation is RelativeKind value
        # When AbsoluteSize RelativeWidth=0, Size.Width=(width 1/100 mm) Width=(width 1/100 mm)
        # When RelativeSize = RelativeWidth=RelativeSize.size, Size.Width and Width become a calculated value base upon RelativeSize.Kind
        super().__init__()
        self._width = width
        self._height = height

    # region Internal Methods

    def _get_rel_width(self, size: RelativeSize) -> int:
        if size.kind == RelativeKind.PAGE:
            doc_size = self.get_page_size()
        else:
            doc_size = self.get_page_text_size()
        percent = size.size / 100
        return round(doc_size.width * percent)

    def _get_rel_height(self, size: RelativeSize) -> int:
        if size.kind == RelativeKind.PAGE:
            doc_size = self.get_page_size()
        else:
            doc_size = self.get_page_text_size()
        percent = size.size / 100
        return round(doc_size.height * percent)

    # endregion Internal Methods

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.text.TextGraphicObject",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies style of current instance.

        Args:
            obj (object): UNO Object that styles are to be applied.
        """
        if kwargs.pop("_apply_clear", True):
            self._clear()
        if self._width is None and self._height is None:
            return
        if self._width:
            if isinstance(self._width, AbsoluteSize):
                self._set(self._props.width, UnitConvert.convert_mm_mm100(self._width.size.value))
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

        if self._height:
            if isinstance(self._height, AbsoluteSize):
                self._set(self._props.height, UnitConvert.convert_mm_mm100(self._height.size.value))
                self._set(self._props.rel_height, 0)

            else:
                rel_height = self._get_rel_height(self._height)
                self._set(self._props.height, rel_height)
                self._set(self._props.rel_height, self._height.size)
                self._set(self._props.rel_height_relation, self._height.kind.value)

        return super().apply(obj, **kwargs)

    # region copy()
    @overload
    def copy(self: _TSize) -> _TSize: ...

    @overload
    def copy(self: _TSize, **kwargs) -> _TSize: ...

    def copy(self: _TSize, **kwargs) -> _TSize:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        if self._width is None:
            cp._width = None
        elif isinstance(self._width, RelativeSize):
            cp._width = dataclasses.replace(self._width)
        else:
            cp._width = self._width.copy()

        if self._height is None:
            cp._height = None
        elif isinstance(self._height, RelativeSize):
            cp._height = dataclasses.replace(self._height)
        else:
            cp._height = self._height.copy()
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
                size_type="",
                width_type="",
            )
        return self._props_internal_attributes

    # endregion Properties
