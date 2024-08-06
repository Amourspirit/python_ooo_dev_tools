"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, overload
from enum import Enum
import math
from ooo.dyn.text.hori_orientation import HoriOrientation
from ooo.dyn.text.vert_orientation import VertOrientation
from ooo.dyn.text.rel_orientation import RelOrientation

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.frame_type_position_props import FrameTypePositionProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_obj import UnitT
from ooodev.utils import props as mProps

_TPosition = TypeVar("_TPosition", bound="Position")


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
    """Vertical Orientation"""

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
    def __init__(self, position: HoriOrient, rel: RelHoriOrient, amount: float | UnitT = 0.0) -> None:
        """
        Constructor

        Args:
            position (HoriOrient): Specifies Horizontal Position.
            rel (RelHoriOrient): Specifies Relative Orientation.
            amount (float, UnitT, optional): Specifies Amount in ``mm`` units or :ref:`proto_unit_obj`. Only effective when position is ``HoriOrient.FROM_LEFT``. Defaults to ``0.0``.
        """

        self._position = position
        self._rel = rel
        try:
            self._amount = amount.get_value_mm()  # type: ignore
        except AttributeError:
            self._amount = float(amount)  # type: ignore

    # endregion Init

    # region Dunder Methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Horizontal):
            result = True
            result = result and self.position == oth.position
            result = result and self.rel == oth.rel
            if not result:
                return False
            return math.isclose(self.amount.value, oth.amount.value, abs_tol=0.02)
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
    def amount(self) -> UnitMM:
        """
        Gets/Sets Amount in ``mm`` units. Only effective when position is ``HoriOrient.FROM_LEFT``.

        Setting also allows ``Unit100MM`` instance.
        """
        return UnitMM(self._amount)

    @amount.setter
    def amount(self, value: float | UnitT):
        try:
            self._amount = value.get_value_mm()  # type: ignore
        except AttributeError:
            self._amount = float(value)  # type: ignore

    # endregion Properties


class Vertical:
    """Vertical Frame Position."""

    # region Init
    def __init__(self, position: VertOrient, rel: RelVertOrient, amount: float | UnitT = 0.0) -> None:
        """
        Constructor

        Args:
            position (VertOrient): Specifies Vertical Position.
            rel (RelVertOrient): Specifies Relative Orientation.
            amount (float, UnitT, optional): Specifies Amount in ``mm`` units or :ref:`proto_unit_obj`.
                Only effective when position is ``VertOrient.FROM_TOP``. Defaults to ``0.0``.
        """

        self._position = position
        self._rel = rel
        try:
            self._amount = amount.get_value_mm()  # type: ignore
        except AttributeError:
            self._amount = float(amount)  # type: ignore

    # endregion Init

    # region Dunder Methods
    def __eq__(self, oth: object) -> bool:
        if isinstance(oth, Vertical):
            result = True
            result = result and self.position == oth.position
            result = result and self.rel == oth.rel
            if not result:
                return False
            return math.isclose(self.amount.value, oth.amount.value, abs_tol=0.02)
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
    def amount(self) -> UnitMM:
        """
        Gets/Sets Amount in ``mm`` units. Only effective when position is ``VertOrient.FROM_TOP``.

        Setting also allows ``Unit100MM`` instance.
        """
        return UnitMM(self._amount)

    @amount.setter
    def amount(self, value: float | UnitT):
        try:
            self._amount = value.get_value_mm()  # type: ignore
        except AttributeError:
            self._amount = float(value)  # type: ignore

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
        keep_boundaries: bool | None = None,
        mirror_even: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            horizontal (Horizontal, optional): Specifies the Horizontal position options. Not used when Anchor
                is set to ``As Character``.
            vertical (Vertical, optional): Specifies the Vertical position options.
            keep_boundaries (bool, optional): Specifies keep inside text boundaries.
            mirror_even (bool, optional): Specifies mirror on even pages. Not used when Anchor is set to
                ``As Character``.
        """
        super().__init__()
        if horizontal is not None:
            self._set_horizontal(horizontal)
        if vertical is not None:
            self._set_vertical(vertical)
        if keep_boundaries is not None:
            self.prop_keep_boundaries = keep_boundaries
        if mirror_even is not None:
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
            self._set(self._props.hori_pos, horizontal.amount.get_value_mm100())
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
            self._set(self._props.vert_pos, vertical.amount.get_value_mm100())
        else:
            self._set(self._props.vert_pos, 0)
        self._set(self._props.vert_orient, vertical.position.value)
        self._set(self._props.vert_rel, vertical.rel.value)

    # endregion Internal Methods

    # region Overrides
    # region copy()
    @overload
    def copy(self: _TPosition) -> _TPosition: ...

    @overload
    def copy(self: _TPosition, **kwargs) -> _TPosition: ...

    def copy(self: _TPosition, **kwargs) -> _TPosition:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._horizontal = None if self._horizontal is None else self._horizontal.copy()
        cp._vertical = None if self._vertical is None else self._vertical.copy()
        return cp

    # endregion copy()
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.Style",
                "com.sun.star.text.BaseFrameProperties",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
                "com.sun.star.text.Shape",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    def _is_valid_obj(self, obj: Any) -> bool:
        """
        Gets if ``obj`` supports one of the services required by style class

        Args:
            obj (Any): UNO object that must have requires service

        Returns:
            bool: ``True`` if has a required service; Otherwise, ``False``
        """
        if super()._is_valid_obj(obj):
            return True
        # check if obj has matching property
        # Some objects such as 'com.sun.star.drawing.shape' sometime support this style.
        # Such is the case when a shape is added to a Writer drawing page.
        # Assume if on attribute matches then they all match.
        return hasattr(obj, self._props.hori_orient)

    # endregion Overrides
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TPosition], obj: Any) -> _TPosition: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TPosition], obj: Any, **kwargs) -> _TPosition: ...

    @classmethod
    def from_obj(cls: Type[_TPosition], obj: Any, **kwargs) -> _TPosition:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Returns:
            Position: Instance that represents Frame Position.
        """

        # pylint: disable=protected-access
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
    def prop_keep_boundaries(self) -> bool | None:
        """Gets/Sets keep inside text boundaries"""
        return self._get(self._props.txt_flow)

    @prop_keep_boundaries.setter
    def prop_keep_boundaries(self, value: bool | None) -> None:
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
    def _props(self) -> FrameTypePositionProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FrameTypePositionProps(
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
