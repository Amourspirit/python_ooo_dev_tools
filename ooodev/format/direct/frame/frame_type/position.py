"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
import dataclasses
from typing import Any, Tuple, cast, Type, TypeVar, overload
from enum import Enum
import math
import uno
from ooo.dyn.text.hori_orientation import HoriOrientation
from ooo.dyn.text.vert_orientation import VertOrientation
from ooo.dyn.text.rel_orientation import RelOrientation

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.unit_convert import UnitConvert
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ...common.props.frame_type_positon_props import FrameTypePositonProps

_TPosition = TypeVar(name="_TPosition", bound="Position")


class HoriOrient(Enum):
    """Horizontal Orientation"""

    LEFT_OR_INSIDE = HoriOrientation.LEFT
    """The object is aligned at the left side. When frame is mirrored the value is ``Inside``; Otherwise, ``Left``."""
    RIGHT_OR_OUTSIDE = HoriOrientation.RIGHT
    """The object is aligned at the right side. When frame is mirrored the value is ``Outside``; Otherwise, ``Right``."""
    CENTER = HoriOrientation.CENTER
    """The object is aligned at the middle."""
    FROM_LEFT_OR_FROM_INSIDE = HoriOrientation.NONE
    """No hard alignment is applied. When frame is mirrored the value is ``From Inside``; Otherwise, ``From Left``."""

    def __int__(self) -> int:
        return self.value


class VertOrient(Enum):
    """Verticial Orientation"""

    TOP = VertOrientation.TOP
    """Aligned at the top"""
    CENTER = VertOrientation.CENTER
    """Aligned at the center"""
    BOTTOM = VertOrientation.BOTTOM
    """Aligned at the bottom"""
    FROM_TOP = VertOrientation.NONE
    """no hard alignment"""


class RelHoriOrient(Enum):
    PARAGRAPH_AREA = RelOrientation.FRAME
    """paragraph, including margins"""
    PARAGRAPH_TEXT_AREA = RelOrientation.PRINT_AREA
    """Paragraph, without margins"""
    LEFT_PARAGRAPH_BORDER = RelOrientation.FRAME_LEFT
    """Inside the left paragraph margin"""
    RIGHT_PARAGRAPH_BORDER = RelOrientation.FRAME_RIGHT
    """Inside the right paragraph margin"""
    LEFT_PAGE_BORDER = RelOrientation.PAGE_LEFT
    """Inside the left page margin"""
    RIGHT_PAGE_BORDER = RelOrientation.PAGE_RIGHT
    """Inside the right page margin"""
    ENTIRE_PAGE = RelOrientation.PAGE_FRAME
    """Page includes margins for page-anchored frames identical with ``PARAGRAPH_AREA``"""
    PAGE_TEXT_AREA = RelOrientation.PAGE_PRINT_AREA
    """Page without borders (for page anchored frames identical with ``PARAGRAPH_TEXT_AREA``)."""


class RelVertOrient(Enum):
    MARGIN = RelOrientation.FRAME
    """paragraph, including margins"""
    PARAGRAPH_TEXT_AREA = RelOrientation.PRINT_AREA
    """Paragraph, without margins"""
    ENTIRE_PAGE = RelOrientation.PAGE_FRAME
    """Page includes margins for page-anchored frames identical with ``PARAGRAPH_AREA``"""
    PAGE_TEXT_AREA = RelOrientation.PAGE_PRINT_AREA
    """Page without borders (for page anchored frames identical with ``PARAGRAPH_TEXT_AREA``)."""


@dataclasses.dataclass(frozen=True)
class Horizontal:
    position: HoriOrient
    """Horizontal Position"""
    rel: RelHoriOrient
    """Relative Orientation"""
    amount: float = 0.0
    """Amount in ``mm`` units. Only effective when position is ``HoriOrient.FROM_LEFT``"""

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Horizontal):
            result = True
            result = result and self.position == oth.position
            result = result and self.rel == oth.rel
            if not result:
                return False
            return math.isclose(self.amount, oth.amount, abs_tol=0.02)
        return NotImplemented


@dataclasses.dataclass(frozen=True)
class Vertical:
    position: VertOrient
    """Vertical Position"""
    rel: RelVertOrient
    """Relative Orientation"""
    amount: float = 0.0
    """Amount in ``mm`` units. Only effective when position is ``VertOrient.FROM_TOP``"""

    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Vertical):
            result = True
            result = result and self.position == oth.position
            result = result and self.rel == oth.rel
            if not result:
                return False
            return math.isclose(self.amount, oth.amount, abs_tol=0.02)
        return NotImplemented


class Position(StyleBase):
    """
    Fill Transparency

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        horizontal: Horizontal | None = None,
        vertical: Vertical | None = None,
        keep_boundries: bool | None = None,
        mirror_even: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            horizontal (Horizontal, optional): Specifies the Horizontal position options.
            vertical (Vertical, optional): Specifies the Vertical position options.
            keep_boundries (bool, optional): Specifies keep inside text boundries.
            mirror_even (bool, optional): Specifies mirror on even pages.
        """
        super().__init__()
        if not horizontal is None:
            self._set_horizontal(horizontal)
        if not vertical is None:
            self._set_vertical(vertical)
        if not keep_boundries is None:
            self.prop_keep_boundries = keep_boundries
        if not mirror_even is None:
            self.prop_mirror_even = mirror_even

        self._horizontal = horizontal
        self._vertical = vertical

    # region Internal Methods
    def _set_horizontal(self, horizontal: Horizontal | None) -> None:
        if horizontal is None:
            self._remove(self._props.hori_orient)
            self._remove(self._props.hori_pos)
            self._remove(self._props.hori_rel)
            return
        if horizontal.position == HoriOrient.FROM_LEFT_OR_FROM_INSIDE:
            self._set(self._props.hori_pos, UnitConvert.convert_mm_mm100(horizontal.amount))
        else:
            self._set(self._props.hori_pos, 0)
        self._set(self._props.hori_orient, horizontal.position.value)
        self._set(self._props.hori_rel, horizontal.rel.value)

    def _set_vertical(self, vertical: Vertical | None) -> None:
        if vertical is None:
            self._remove(self._props.vert_orient)
            self._remove(self._props.vert_pos)
            self._remove(self._props.vert_rel)
            return
        if vertical.position == VertOrient.FROM_TOP:
            self._set(self._props.vert_pos, UnitConvert.convert_mm_mm100(vertical.amount))
        else:
            self._set(self._props.vert_pos, 0)
        self._set(self._props.vert_orient, vertical.position.value)
        self._set(self._props.vert_rel, vertical.rel.value)

    # endregion Internal Methods

    # region Overrides
    def copy(self: _TPosition) -> _TPosition:
        cp = super().copy()
        if self._horizontal is None:
            cp._horizontal = None
        else:
            cp._horizontal = dataclasses.replace(self._horizontal)
        if self._vertical is None:
            cp._vertical = None
        else:
            cp._vertical = dataclasses.replace(self._vertical)
        return cp

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

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TPosition], obj: object) -> _TPosition:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TPosition], obj: object, **kwargs) -> _TPosition:
        ...

    @classmethod
    def from_obj(cls: Type[_TPosition], obj: object, **kwargs) -> _TPosition:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Returns:
            Position: Instance that represents Frame Position.
        """

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        pp = inst._props
        hori_orient = int(mProps.Props.get(obj, pp.hori_orient, HoriOrient.CENTER.value))
        inst._set(pp.hori_orient, hori_orient)

        hori_pos = int(mProps.Props.get(obj, pp.hori_pos, 0))
        inst._set(pp.hori_pos, hori_pos)

        hori_rel = int(mProps.Props.get(obj, pp.hori_rel, RelHoriOrient.PARAGRAPH_TEXT_AREA.value))
        inst._set(pp.hori_rel, hori_rel)

        vert_orient = int(mProps.Props.get(obj, pp.vert_orient, VertOrient.TOP.value))
        inst._set(pp.vert_orient, vert_orient)

        vert_pos = int(mProps.Props.get(obj, pp.vert_pos, 0))
        inst._set(pp.vert_pos, vert_pos)

        vert_rel = int(mProps.Props.get(obj, pp.vert_rel, RelVertOrient.PARAGRAPH_TEXT_AREA.value))
        inst._set(pp.hori_rel, vert_rel)

        txt_flow = bool(mProps.Props.get(obj, pp.txt_flow, False))
        inst._set(pp.txt_flow, txt_flow)

        page_toggle = bool(mProps.Props.get(obj, pp.page_toggle, False))
        inst._set(pp.page_toggle, page_toggle)

        inst._horizontal = Horizontal(
            position=HoriOrient(hori_orient),
            rel=RelHoriOrient(hori_rel),
            amount=UnitConvert.convert_mm100_mm(hori_pos),
        )
        inst._vertical = Vertical(
            position=VertOrient(vert_orient),
            rel=RelVertOrient(vert_rel),
            amount=UnitConvert.convert_mm100_mm(vert_pos),
        )

        return inst

    # endregion from_obj()
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_keep_boundries(self) -> bool | None:
        """Gets/Sets keep inside text boundries"""
        return self._get(self._props.txt_flow)

    @prop_keep_boundries.setter
    def prop_keep_boundries(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.txt_flow)
            return
        self._set(self._props.txt_flow, value)

    @property
    def prop_mirror_even(self) -> bool | None:
        """Gets/Sets Mirror on even pages"""
        return self._get(self._props.page_toggle)

    @prop_mirror_even.setter
    def prop_mirror_even(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.page_toggle)
            return
        self._set(self._props.page_toggle, value)

    @property
    def prop_horizontal(self) -> HoriOrient | None:
        """Gets/Sets horizontal value"""
        return self._horizontal

    @prop_horizontal.setter
    def prop_horizontal(self, value: Horizontal | None) -> None:
        self._set_horizontal(value)

    @property
    def prop_vertical(self) -> HoriOrient | None:
        """Gets/Sets vertical value"""
        return self._vertical

    @prop_vertical.setter
    def prop_vertical(self, value: Vertical | None) -> None:
        self._set_vertical(value)

    @property
    def _props(self) -> FrameTypePositonProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FrameTypePositonProps(
                hori_orient="HoriOrient",
                hori_pos="HoriOrientPosition",
                hori_rel="HoriOrientRelation",
                vert_orient="VertOrient",
                vert_pos="VertOrientPosition",
                vert_rel="VertOrientRelation",
                txt_flow="IsFollowingTextFlow",
                page_toggle="PageToggle",
            )

        return self._props_internal_attributes
