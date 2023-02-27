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
from .....utils.data_type.unit_100_mm import Unit100MM
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
    FROM_LEFT_OR_INSIDE = HoriOrientation.NONE
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
    FROM_TOP_OR_BOTTOM = VertOrientation.NONE
    """No hard alignment. From Top is only used when Anchor is ``To character``. From Bottom is only used when Anchor is ``To character`` or ``As character``."""
    BELOW_CHAR = VertOrientation.CHAR_BOTTOM
    """Aligned at the bottom of a character (anchored to character). Only used when Anchor is set ``To character``."""


class RelHoriOrient(Enum):
    # when Anchor is set to page, the paragraph enums are not present in the Writer Dialog.
    # when Anchor is set to character.
    PARAGRAPH_AREA = RelOrientation.FRAME
    """Paragraph, including margins. Only used when Anchor is set ``To paragraph`` or ``To Character``."""
    PARAGRAPH_TEXT_AREA = RelOrientation.PRINT_AREA
    """Paragraph, without margins. Only used when Anchor is set ``To paragraph`` or ``To Character``."""
    LEFT_PARAGRAPH_BORDER = RelOrientation.FRAME_LEFT
    """Inside the left paragraph margin. Only used when Anchor is set ``To paragraph`` or ``To Character``."""
    RIGHT_PARAGRAPH_BORDER = RelOrientation.FRAME_RIGHT
    """Inside the right paragraph margin. Only used when Anchor is set ``To paragraph`` or ``To Character``."""
    LEFT_PAGE_BORDER = RelOrientation.PAGE_LEFT
    """Inside the left page margin."""
    RIGHT_PAGE_BORDER = RelOrientation.PAGE_RIGHT
    """Inside the right page margin"""
    ENTIRE_PAGE = RelOrientation.PAGE_FRAME
    """Page includes margins for page-anchored frames identical with ``PARAGRAPH_AREA``."""
    PAGE_TEXT_AREA = RelOrientation.PAGE_PRINT_AREA
    """Page without borders (for page anchored frames identical with ``PARAGRAPH_TEXT_AREA``)."""
    CHARACTER = RelOrientation.CHAR
    """As character. Only used when Anchor is set ``To character``."""


class RelVertOrient(Enum):
    MARGIN = RelOrientation.FRAME
    """Paragraph, including margins. Only used when Anchor is set ``To paragraph`` or ``To character``."""
    PARAGRAPH_TEXT_AREA_OR_BASE_LINE = RelOrientation.PRINT_AREA
    """Paragraph, without margins. Only used when Anchor is set ``To paragraph``, ``To character`` or ``As character``."""
    ENTIRE_PAGE_OR_ROW = RelOrientation.PAGE_FRAME
    """Page includes margins for page-anchored frames identical with ``PARAGRAPH_AREA``. Only used when Anchor is set ``To page``, ``To paragraph``, ``To character`` or ``As character``."""
    PAGE_TEXT_AREA = RelOrientation.PAGE_PRINT_AREA
    """Page without borders (for page anchored frames identical with ``PARAGRAPH_TEXT_AREA``). Only used when Anchor is set ``To page``, ``To paragraph`` or ``To character``."""
    CHARACTER = RelOrientation.CHAR
    """As character. Only used when Anchor is set ``To character``."""
    LINE_OF_TEXT = RelOrientation.TEXT_LINE
    """at the top of the text line, only sensible for vertical orientation. Only used when Anchor ``To character`` and ``VertOrient`` is ``TOP``, ``BOTTOM``, ``CENTER`` or ``FROM_TOP_OR_BOTTOM``."""


class Horizontal:
    """Horizontal Frame Position. Not used when Anchor is set to ``As Character``."""

    # region Init
    def __init__(self, position: HoriOrient, rel: RelHoriOrient, amount: float | Unit100MM = 0.0) -> None:
        """
        Constructor

        Args:
            position (HoriOrient): Specifies Horizontal Position.
            rel (RelHoriOrient): Specifies Relative Orientation.
            amount (float, Unit100MM, optional): Spedifies Amount in ``mm`` units or ``1/100th mm`` units. Only effective when position is ``HoriOrient.FROM_LEFT``. Defaults to ``0.0``.
        """

        self._position = position
        self._rel = rel
        if isinstance(amount, Unit100MM):
            self._amount = amount.get_value_mm()
        else:
            self._amount = amount

    # endregion Init

    # region Dunder Methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Horizontal):
            result = True
            result = result and self.position == oth.position
            result = result and self.rel == oth.rel
            if not result:
                return False
            return math.isclose(self.amount, oth.amount, abs_tol=0.02)
        return NotImplemented

    # endregion Dunder Methods

    # region Methods
    def copy(self) -> Horizontal:
        """Gets a copy of instance as a new instance"""
        inst = super(Horizontal, self.__class__).__new__(self.__class__)
        inst.__init__(position=self.position, rel=self.rel, amount=self.amount)
        return inst

    # endregion Methods

    # region Properties
    @property
    def position(self) -> HoriOrient:
        """
        Gets/Sets Horizontal Position.
        """
        return self._position

    @position.setter
    def position(self, value: HoriOrient):
        self._position = value

    @property
    def rel(self) -> RelHoriOrient:
        """
        Gets/Sets Relative Orientation
        """
        return self._rel

    @rel.setter
    def rel(self, value: RelHoriOrient):
        self._rel = value

    @property
    def amount(self) -> float:
        """
        Gets/Sets Amount in ``mm`` units. Only effective when position is ``HoriOrient.FROM_LEFT``.

        Setting also allows ``Unit100MM`` instance.
        """
        return self._amount

    @amount.setter
    def amount(self, value: float | Unit100MM):
        if isinstance(value, Unit100MM):
            self._amount = value.get_value_mm()
        else:
            self._amount = value

    # endregion Properties


class Vertical:
    """Vertical Frame Position."""

    # region Init
    def __init__(self, position: VertOrient, rel: RelVertOrient, amount: float | Unit100MM = 0.0) -> None:
        """
        Constructor

        Args:
            position (VertOrient): Specifies Vertical Position.
            rel (RelVertOrient): Specifies Relative Orientation.
            amount (float, Unit100MM, optional): Spedifies Amount in ``mm`` units or ``1/100th mm`` units. Only effective when position is ``VertOrient.FROM_TOP``. Defaults to ``0.0``.
        """

        self._position = position
        self._rel = rel
        if isinstance(amount, Unit100MM):
            self._amount = amount.get_value_mm()
        else:
            self._amount = amount

    # endregion Init

    # region Dunder Methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Vertical):
            result = True
            result = result and self.position == oth.position
            result = result and self.rel == oth.rel
            if not result:
                return False
            return math.isclose(self.amount, oth.amount, abs_tol=0.02)
        return NotImplemented

    # endregion Dunder Methods

    # region Methods
    def copy(self) -> Vertical:
        """Gets a copy of instance as a new instance"""
        inst = super(Vertical, self.__class__).__new__(self.__class__)
        inst.__init__(position=self.position, rel=self.rel, amount=self.amount)
        return inst

    # endregion Methods

    # region Properties
    @property
    def position(self) -> VertOrient:
        """
        Gets/Sets Vertical Position.
        """
        return self._position

    @position.setter
    def position(self, value: VertOrient):
        self._position = value

    @property
    def rel(self) -> RelVertOrient:
        """
        Gets/Sets Relative Orientation
        """
        return self._rel

    @rel.setter
    def rel(self, value: RelVertOrient):
        self._rel = value

    @property
    def amount(self) -> float:
        """
        Gets/Sets Amount in ``mm`` units. Only effective when position is ``VertOrient.FROM_TOP``.

        Setting also allows ``Unit100MM`` instance.
        """
        return self._amount

    @amount.setter
    def amount(self, value: float | Unit100MM):
        if isinstance(value, Unit100MM):
            self._amount = value.get_value_mm()
        else:
            self._amount = value

    # endregion Properties


class Position(StyleBase):
    """
    Fill Transparency

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        horizontal: Horizontal | None = None,
        vertical: Vertical | None = None,
        keep_boundries: bool | None = None,
        mirror_even: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            horizontal (Horizontal, optional): Specifies the Horizontal position options. Not used when Anchor is set to ``As Character``.
            vertical (Vertical, optional): Specifies the Vertical position options.
            keep_boundries (bool, optional): Specifies keep inside text boundries.
            mirror_even (bool, optional): Specifies mirror on even pages. Not used when Anchor is set to ``As Character``.
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
        if horizontal.position == HoriOrient.FROM_LEFT_OR_INSIDE:
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
        if vertical.position == VertOrient.FROM_TOP_OR_BOTTOM:
            self._set(self._props.vert_pos, UnitConvert.convert_mm_mm100(vertical.amount))
        else:
            self._set(self._props.vert_pos, 0)
        self._set(self._props.vert_orient, vertical.position.value)
        self._set(self._props.vert_rel, vertical.rel.value)

    # endregion Internal Methods

    # region Overrides
    # region copy()
    @overload
    def copy(self: _TPosition) -> _TPosition:
        ...

    @overload
    def copy(self: _TPosition, **kwargs) -> _TPosition:
        ...

    def copy(self: _TPosition, **kwargs) -> _TPosition:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        if self._horizontal is None:
            cp._horizontal = None
        else:
            # cp._horizontal = dataclasses.replace(self._horizontal)
            cp._horizontal = self._horizontal.copy()
        if self._vertical is None:
            cp._vertical = None
        else:
            # cp._vertical = dataclasses.replace(self._vertical)
            cp._vertical = self._vertical.copy()
        return cp

    # endregion copy()
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.Style",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
                "com.sun.star.text.BaseFrame",
            )
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

        vert_rel = int(mProps.Props.get(obj, pp.vert_rel, RelVertOrient.PARAGRAPH_TEXT_AREA_OR_BASE_LINE.value))
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
    def prop_horizontal(self) -> Horizontal | None:
        """Gets/Sets horizontal value"""
        return self._horizontal

    @prop_horizontal.setter
    def prop_horizontal(self, value: Horizontal | None) -> None:
        self._set_horizontal(value)

    @property
    def prop_vertical(self) -> Vertical | None:
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
