# region Imports
from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, NamedTuple, Union, cast, overload
from enum import Enum
import math
from ooo.dyn.text.hori_orientation import HoriOrientation

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.data_type.size import Size
from ooodev.units.unit_obj import UnitT
from ooodev.units.unit_mm import UnitMM
from ooodev.format.inner.common.abstract.abstract_document import AbstractDocument
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.common.props.table_properties_props import TablePropertiesProps

from ooodev.units.unit_convert import UnitConvert

# endregion Imports

# region Types

_TTableProperties = TypeVar(name="_TTableProperties", bound="TableProperties")
_TSharedAuto = TypeVar(name="_TSharedAuto", bound="_SharedAuto")
_TTblAuto = TypeVar(name="_TTblAuto", bound="_TblAuto")
_TTblRelLeftByWidth = TypeVar(name="_TTblRelLeftByWidth", bound="_TblRelLeftByWidth")
TblAbsUnit = Union[float, UnitT]
TblRelUnit = Union[int, Intensity]

# endregion Types


# region Enums
class TableAlignKind(Enum):
    MANUAL = HoriOrientation.NONE
    """No hard alignment is applied."""
    AUTO = HoriOrientation.FULL
    """The object uses the full space (for text tables only)."""
    CENTER = HoriOrientation.CENTER
    """The object is aligned at the middle."""
    LEFT = HoriOrientation.LEFT
    """The object is aligned at the left side."""
    RIGHT = HoriOrientation.RIGHT
    """The object is aligned at the right side."""
    FROM_LEFT = HoriOrientation.LEFT_AND_WIDTH
    """
    The left offset and the width of the object are defined.
    
    For text tables this means that the tables position is defined by the left margin and the width.
    """


# endregion Enums


# region Tuples
class _RelVals(NamedTuple):
    """Relative values. Total expected to add up to ``100``"""

    left: int
    right: int
    balance: int


# endregion Tuples

# region Module Methods


def _get_default_tbl_props() -> TablePropertiesProps:
    return TablePropertiesProps(
        name="Name",
        width="Width",
        left="LeftMargin",
        top="TopMargin",
        right="RightMargin",
        bottom="BottomMargin",
        is_rel="IsWidthRelative",
        rel_width="RelativeWidth",
        hori_orient="HoriOrient",
    )


def _get_default_tbl_services() -> Tuple[str]:
    return ("com.sun.star.text.TextTable",)


# endregion Module Methods


class _SharedAuto(AbstractDocument):
    """
    Automatically set table width.
    """

    # region Init
    def __init__(
        self,
        *,
        width: TblAbsUnit | TblRelUnit | None = None,
        left: TblAbsUnit | TblRelUnit | None = None,
        right: TblAbsUnit | TblRelUnit | None = None,
        above: TblAbsUnit | TblRelUnit | None = None,
        below: TblAbsUnit | TblRelUnit | None = None,
    ) -> None:
        """
        Constructor

        Args:
            width (TblAbsUnit, optional): Table width value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            left (TblAbsUnit, optional): Spacing Left value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            right (TblAbsUnit, optional): Spacing Right value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            above (TblAbsUnit, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (TblAbsUnit, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        # only Above and Below properties are used in this class
        super().__init__()
        self.prop_width = width
        self.prop_left = left
        self.prop_right = right
        if above is not None:
            self.prop_above = above
        if below is not None:
            self.prop_below = below
        self._post_init()

    # endregion Init

    # region internal methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.AUTO.value)

    def _get_relative_values(self) -> _RelVals:
        left100: int = self._get(self._props.left)
        right100: int = self._get(self._props.right)
        page_txt_width = self._prop_page_text_size.width
        left = 0 if left100 == 0 else round((left100 / page_txt_width) * 100)
        right = 0 if right100 == 0 else round((right100 / page_txt_width) * 100)
        width = 100 - (left + right)
        return _RelVals(left=left, right=right, balance=width)

    def _set_props_from_obj(self, obj: Any) -> None:
        if not self._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{self.__class__.__name__}"')
        self._set(self._props.left, mProps.Props.get(obj, self._props.left))
        self._set(self._props.right, mProps.Props.get(obj, self._props.right))
        self._set(self._props.top, mProps.Props.get(obj, self._props.top))
        self._set(self._props.bottom, mProps.Props.get(obj, self._props.bottom))
        self._set(self._props.width, mProps.Props.get(obj, self._props.width))
        self._set(self._props.hori_orient, mProps.Props.get(obj, self._props.hori_orient))
        self._set(self._props.is_rel, mProps.Props.get(obj, self._props.is_rel))
        self._set(self._props.rel_width, mProps.Props.get(obj, self._props.rel_width))
        self._prop_width = int(self._get(self._props.width))
        self._prop_left = int(self._get(self._props.left))
        self._prop_right = int(self._get(self._props.right))

    def _get_prop_width(self) -> UnitMM | None:
        if self._prop_width is None:
            return None
        return UnitMM.from_mm100(self._prop_width)

    def _set_prop_width(self, value: TblAbsUnit | TblRelUnit | None) -> None:
        if value is None:
            self._prop_width = None
            return
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)  # type: ignore
        # min table width seems to be in the area of 1.22 mm
        val = max(val, 122)
        self._prop_width = val

    def _get_prop_left(self) -> UnitMM | None:
        return None if self._prop_left is None else UnitMM.from_mm100(self._prop_left)

    def _set_prop_left(self, value: TblAbsUnit | TblRelUnit | None) -> None:
        if value is None:
            self._prop_left = None
            return
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)  # type: ignore
        self._prop_left = val

    def _get_prop_right(self) -> UnitMM | None:
        if self._prop_right is None:
            return None
        return UnitMM.from_mm100(self._prop_right)

    def _set_prop_right(self, value: TblAbsUnit | TblRelUnit | None) -> None:
        if value is None:
            self._prop_right = None
            return
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)  # type: ignore
        self._prop_right = val

    def _get_prop_above(self) -> UnitMM | None:
        pv = self._get(self._props.top)
        return None if pv is None else UnitMM.from_mm100(pv)

    def _set_prop_above(self, value: TblAbsUnit | TblRelUnit | None) -> None:
        if value is None:
            self._remove(self._props.top)
            return
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)  # type: ignore
        self._set(self._props.top, val)

    def _get_prop_below(self) -> UnitMM | None:
        pv = self._get(self._props.bottom)
        return None if pv is None else UnitMM.from_mm100(pv)

    def _set_prop_below(self, value: TblAbsUnit | TblRelUnit | None) -> None:
        if value is None:
            self._remove(self._props.bottom)
            return
        try:
            val = value.get_value_mm100()  # type: ignore
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)  # type: ignore
        self._set(self._props.bottom, val)

    # endregion internal methods

    # region overrides
    def copy(self: _TSharedAuto, **kwargs) -> _TSharedAuto:
        """Gets a copy of the instance"""
        cp = super().copy(**kwargs)
        cp._prop_width = self._prop_width
        cp._prop_left = self._prop_left
        cp._prop_right = self._prop_right
        return cp

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = _get_default_tbl_services()
        return self._supported_services_values

    def _set_defaults(self) -> None:
        if not self._has(self._props.width):
            self._set(self._props.width, self._prop_page_text_size.width)
            self._prop_width = self._prop_page_text_size.width
        if not self._has(self._props.left):
            self._set(self._props.left, 0)
            self._prop_left = 0
        if not self._has(self._props.right):
            self._set(self._props.right, 0)
            self._prop_right = 0
        if not self._has(self._props.rel_width):
            self._set(self._props.rel_width, 0)

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies Styling to object

        Args:
            obj (object): UNO Table Object

        Returns:
            None:
        """
        self._set_defaults()
        super().apply(obj, **kwargs)
        try:
            self._set_props_from_obj(obj)
        except Exception as e:
            mLo.Lo.print(f"Unable to set property values for {self.__class__.__name__} after apply method.")
            mLo.Lo.print(f"  {e}")

    # endregion overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSharedAuto], obj: Any) -> _TSharedAuto: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSharedAuto], obj: Any, **kwargs) -> _TSharedAuto: ...

    @classmethod
    def from_obj(cls: Type[_TSharedAuto], obj: Any, **kwargs) -> _TSharedAuto:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblAuto: ``TblAuto`` Instance.
        """
        return _SharedAuto._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TSharedAuto], obj: Any, **kwargs) -> _TSharedAuto:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Options: Instance that represents Frame Wrap Settings.
        """
        # pylint: disable=protected-access
        inst = clazz(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{clazz.__name__}"')
        inst._set_props_from_obj(obj)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()
    # endregion static methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.TABLE
        return self._format_kind_prop

    @property
    def prop_width(self) -> UnitMM | None:
        """
        Gets/Sets Width.
        """
        return self._get_prop_width()

    @prop_width.setter
    def prop_width(self, value: TblAbsUnit | TblRelUnit | None):
        self._set_prop_width(value)

    @property
    def prop_left(self) -> UnitMM | None:
        """
        Gets/Sets Left.
        """
        return self._get_prop_left()

    @prop_left.setter
    def prop_left(self, value: TblAbsUnit | TblRelUnit | None):
        self._set_prop_left(value)

    @property
    def prop_right(self) -> UnitMM | None:
        """
        Gets/Sets Right.
        """
        return self._get_prop_right()

    @prop_right.setter
    def prop_right(self, value: TblAbsUnit | TblRelUnit | None):
        self._set_prop_right(value)

    @property
    def prop_above(self) -> UnitMM | None:
        """
        Gets/Sets above.
        """
        return self._get_prop_above()

    @prop_above.setter
    def prop_above(self, value: TblAbsUnit | TblRelUnit | None):
        self._set_prop_above(value)

    @property
    def prop_below(self) -> UnitMM | None:
        """
        Gets/Sets below.
        """
        return self._get_prop_below()

    @prop_below.setter
    def prop_below(self, value: TblAbsUnit | TblRelUnit | None):
        self._set_prop_below(value)

    @property
    def _props(self) -> TablePropertiesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _get_default_tbl_props()
        return self._props_internal_attributes

    @property
    def _prop_page_text_size(self) -> Size:
        """Page Text Size in ``1/100th mm``: Only Call In ``apply()`` Method"""
        # keep private. Not available until writer doc is created.
        try:
            return self._prop_page_text_size_attrib
        except AttributeError:
            self._prop_page_text_size_attrib = self.get_page_text_size()
        return self._prop_page_text_size_attrib

    @property
    def _prop_page_size(self) -> Size:
        """Page Size in ``1/100th mm``: Only Call In ``apply()`` Method"""
        # keep private. Not available until writer doc is created.
        try:
            return self._prop_page_size_attrib
        except AttributeError:
            self._prop_page_size_attrib = self.get_page_size()
        return self._prop_page_size_attrib

    # endregion Properties


# region Table size in MM units


class _TblAuto(_SharedAuto):
    def __init__(
        self,
        *,
        width: TblAbsUnit | None = None,
        left: TblAbsUnit | None = None,
        right: TblAbsUnit | None = None,
        above: TblAbsUnit | None = None,
        below: TblAbsUnit | None = None,
    ) -> None:
        """
        Constructor

        Args:
            width (TblAbsUnit, optional): Table width value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            left (TblAbsUnit, optional): Spacing Left value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            right (TblAbsUnit, optional): Spacing Right value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            above (TblAbsUnit, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (TblAbsUnit, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        # only Above and Below properties are used in this class
        super().__init__(width=width, left=left, right=right, above=above, below=below)

    # region Static Methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTblAuto], obj: Any) -> _TTblAuto: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto: ...

    @classmethod
    def from_obj(cls: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblAuto: ``TblAuto`` Instance.
        """
        return _TblAuto._from_obj(cls, obj, **kwargs)

    # endregion from_obj()

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Options: Instance that represents Frame Wrap Settings.
        """
        inst = clazz(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{clazz.__name__}"')
        inst._set_props_from_obj(obj)
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_width(self) -> UnitMM | None:
        """
        Gets/Sets Width.
        """
        return self._get_prop_width()

    @prop_width.setter
    def prop_width(self, value: TblAbsUnit | None):
        self._set_prop_width(value)

    @property
    def prop_left(self) -> UnitMM | None:
        """
        Gets/Sets Left.
        """
        return self._get_prop_left()

    @prop_left.setter
    def prop_left(self, value: TblAbsUnit | None):
        self._set_prop_left(value)

    @property
    def prop_right(self) -> UnitMM | None:
        """
        Gets/Sets Right.
        """
        return self._get_prop_right()

    @prop_right.setter
    def prop_right(self, value: TblAbsUnit | None):
        self._set_prop_right(value)

    @property
    def prop_above(self) -> UnitMM | None:
        """
        Gets/Sets above.
        """
        return self._get_prop_above()

    @prop_above.setter
    def prop_above(self, value: TblAbsUnit | None):
        self._set_prop_above(value)

    @property
    def prop_below(self) -> UnitMM | None:
        """
        Gets/Sets below.
        """
        return self._get_prop_below()

    @prop_below.setter
    def prop_below(self, value: TblAbsUnit | None):
        self._set_prop_below(value)

    # endregion Properties


class _TblCenterWidth(_TblAuto):
    """
    Sets the table width using a width value and centers the table.
    """

    # Only width, above and below properties are used in this class
    # only width is used directly

    # region internal Methods
    def _set_width_properties(self, obj: Any) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
        if self.prop_width is None:
            return
        page_txt_width = self._prop_page_text_size.width
        width100 = self.prop_width.get_value_mm100()
        if width100 > page_txt_width * 2:
            width100 = page_txt_width * 2
            margin = -page_txt_width
        else:
            # calculate the margins
            # if width100 is greater then page_txt_width then margins are negative.
            margin = page_txt_width - width100

        if margin == 0:
            left_margin = 0
            right_margin = 0
        else:
            left_margin = round(margin / 2)
            right_margin = left_margin

        if margin > 0:
            while right_margin + left_margin > margin:
                # just in case rounding add a digit or two
                right_margin -= 1
        elif margin < 0:
            while right_margin + left_margin < margin:
                # just in case rounding add a digit or two
                right_margin += 1

        self._set(self._props.left, left_margin)
        self._set(self._props.right, right_margin)
        self._set(self._props.width, width100)

    # endregion internal Methods

    # region overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.CENTER.value)

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies Styling to object

        Args:
            obj (object): UNO Table Object

        Returns:
            None:
        """
        self._set_width_properties(obj)
        super().apply(obj, **kwargs)

    # endregion overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblCenterWidth: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblCenterWidth: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblCenterWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblCenterWidth: ``TblCenterWidth`` Instance.
        """
        return _TblCenterWidth._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblAuto._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblLeft(_TblCenterWidth):
    """
    Sets the right table margin.
    """

    # right, above, and below properties are used in this class.
    # only right is used directly in this class.

    # region internal Methods

    def _set_width_properties(self, obj: Any) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
        if self.prop_right is None:
            return
        self._set(self._props.left, 0)
        page_txt_width = self._prop_page_text_size.width
        margin = self.prop_right.get_value_mm100()

        if margin == 0:
            self._set(self._props.right, 0)
            self._set(self._props.width, page_txt_width)
            return

        if margin > 0:
            # minimum width allowed for table seems to be 1.22 mm
            min_width100 = 122
            if margin > page_txt_width + min_width100:
                margin = page_txt_width - min_width100
            width100 = page_txt_width - margin
            self._set(self._props.width, width100)
            self._set(self._props.right, margin)
            return

        # margin is negative

        if abs(margin) > page_txt_width:
            margin = -page_txt_width

        width100 = abs(margin) + page_txt_width
        self._set(self._props.width, width100)
        self._set(self._props.right, margin)

    # endregion internal Methods

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.LEFT.value)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblLeft: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblLeft: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblLeft:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblLeft: ``TblLeft`` Instance.
        """
        return _TblLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblCenterWidth._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblLeftWidth(_TblLeft):
    """
    Sets the right table margin.
    """

    # right, above, and below properties are used in this class.
    # only right is used directly in this class.

    # region internal Methods

    def _set_width_properties(self, obj: Any) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
        if self.prop_width is None:
            return
        self._set(self._props.left, 0)
        tbl_min_width = 122
        page_txt_width = self._prop_page_text_size.width
        width = self.prop_width.get_value_mm100()
        if width <= tbl_min_width:
            self._set(self._props.right, page_txt_width - tbl_min_width)
            self._set(self._props.width, tbl_min_width)
            return

        if width >= page_txt_width:
            self._set(self._props.right, 0)
            self._set(self._props.width, page_txt_width)
            return

        right = page_txt_width - width
        self._set(self._props.right, right)
        self._set(self._props.width, width)

    # endregion internal Methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblLeftWidth: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblLeftWidth: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblLeftWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblLeftWidth: ``TblLeftWidth`` Instance.
        """
        return _TblLeftWidth._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblLeft._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblRight(_TblLeft):
    """
    Sets the left table margin.
    """

    # left, above, and below properties are used in this class.
    # only left is used directly in this class.

    # region override Methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.RIGHT.value)

    def _set_width_properties(self, obj: Any) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
        if self.prop_left is None:
            return
        self._set(self._props.right, 0)
        page_txt_width = self._prop_page_text_size.width
        margin = self.prop_left.get_value_mm100()

        if margin == 0:
            self._set(self._props.left, 0)
            self._set(self._props.width, page_txt_width)
            return

        if margin > 0:
            # minimum width allowed for table seems to be 1.22 mm
            min_width100 = 122
            if margin > page_txt_width + min_width100:
                margin = page_txt_width - min_width100
            width100 = page_txt_width - margin
            self._set(self._props.width, width100)
            self._set(self._props.left, margin)
            return

        # margin is negative

        if abs(margin) > page_txt_width:
            margin = -page_txt_width

        width100 = abs(margin) + page_txt_width
        self._set(self._props.width, width100)
        self._set(self._props.left, margin)

    # endregion override Methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblRight: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRight: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRight:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRight: ``TblRight`` Instance.
        """
        return _TblRight._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblLeft._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblRightWidth(_TblRight):
    """
    Sets the left table margin.
    """

    # region override Methods
    def _set_width_properties(self, obj: Any) -> None:
        if self.prop_width is None:
            return
        self._set(self._props.right, 0)
        tbl_min_width = 122
        page_txt_width = self._prop_page_text_size.width
        width = self.prop_width.get_value_mm100()
        if width <= tbl_min_width:
            self._set(self._props.left, page_txt_width - tbl_min_width)
            self._set(self._props.width, tbl_min_width)
            return

        if width >= page_txt_width:
            self._set(self._props.left, 0)
            self._set(self._props.width, page_txt_width)
            return

        left = page_txt_width - width
        self._set(self._props.left, left)
        self._set(self._props.width, width)

    # endregion override Methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblRightWidth: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRightWidth: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRightWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRightWidth: ``TblRightWidth`` Instance.
        """
        return _TblRightWidth._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblRight._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblCenterLeft(_TblLeft):
    """
    Sets the table width using a width value and centers the table.
    """

    # left, above, and below properties are used in this class.
    # only left is used directly in this class.

    # region override Methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.CENTER.value)

    def _set_width_properties(self, obj: Any) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
        if self.prop_left is None:
            return
        page_txt_width = self._prop_page_text_size.width
        margin = self.prop_left.get_value_mm100()
        # minimum width allowed for table seems to be 1.22 mm
        min_width100 = 122

        if margin == 0:
            self._set(self._props.left, 0)
            self._set(self._props.right, 0)
            self._set(self._props.width, page_txt_width)
            return

        # left_margin = round(abs(margin) / 2)
        width = page_txt_width - (margin * 2)
        if width < min_width100:
            self._set(self._props.left, round(page_txt_width - (min_width100 / 2)))
            self._set(self._props.right, round(page_txt_width - (min_width100 / 2)))
            self._set(self._props.width, min_width100)
            return

        self._set(self._props.width, width)
        self._set(self._props.left, margin)
        self._set(self._props.right, margin)

    # endregion override Methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblCenterLeft: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblCenterLeft: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblCenterLeft:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblCenterLeft: ``TblCenterLeft`` Instance.
        """
        return _TblCenterLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblLeft._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblFromLeft(_TblLeft):
    """
    Sets the table width from left using a margin value.
    """

    # left, above, and below properties are used in this class.
    # only left is used directly in this class.

    # region override Methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.FROM_LEFT.value)

    def _set_width_properties(self, obj: Any) -> None:
        # the total of width + left + right = page_txt_width
        # when left changes the width remains the same and right can become a neg value
        # if width exceeds page_txt_width then is set to page_txt_width, left and right become 0
        # min table width 1.22 mm

        # if left exceeds page_txt_width then left is set to page_txt_width, width is unchanged and right is set to neg width
        if self.prop_left is None:
            return
        page_txt_width = self._prop_page_text_size.width
        margin = self.prop_left.get_value_mm100()
        tbl_width = int(mProps.Props.get(obj, self._props.width))
        if margin >= page_txt_width:
            self._set(self._props.left, page_txt_width)
            self._set(self._props.right, -tbl_width)
            return

        right = page_txt_width - tbl_width - margin
        assert right + margin + tbl_width == page_txt_width
        self._set(self._props.left, margin)
        self._set(self._props.right, right)

    # endregion override Methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblFromLeft: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblFromLeft: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblFromLeft:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblFromLeft: ``TblFromLeft`` Instance.
        """
        return _TblFromLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblLeft._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblFromLeftWidth(_TblCenterWidth):
    """
    Sets the table width from left using a width value.
    """

    # width, above, and below properties are used in this class.
    # only width is used directly in this class.

    # region override Methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.FROM_LEFT.value)

    def _set_width_properties(self, obj: Any) -> None:
        # sourcery skip: extract-duplicate-method
        # the total of width + left + right = page_txt_width.
        # when width changes left stays the same and right gets adjusted conditionally.
        #   if right is 0 then width will reduce left.
        # if width exceeds page_txt_width then is set to page_txt_width, left and right become 0/
        # min table width 1.22 mm.

        if self.prop_width is None:
            return
        page_txt_width = self._prop_page_text_size.width
        tbl_min_width = 122

        width100 = self.prop_width.get_value_mm100()
        if width100 >= page_txt_width:
            self._set(self._props.right, 0)
            self._set(self._props.left, 0)
            self._set(self._props.width, page_txt_width)
            return

        if width100 < tbl_min_width:  # min width:
            self._set(self._props.right, 0)
            self._set(self._props.left, page_txt_width - tbl_min_width)
            self._set(self._props.width, tbl_min_width)
            return

        left = page_txt_width - width100

        self._set(self._props.right, 0)
        self._set(self._props.left, left)
        self._set(self._props.width, width100)

    # endregion override Methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblFromLeftWidth: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblFromLeftWidth: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblFromLeftWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblFromLeftWidth: ``TblFromLeftWidth`` Instance.
        """
        return _TblFromLeftWidth._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblCenterWidth._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblManualLeftRight(_TblCenterWidth):
    """
    Sets the table width manually.
    """

    # all properties of init are used in this class

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.MANUAL.value)

    def _set_width_properties(self, obj: Any) -> None:
        # sourcery skip: extract-duplicate-method, use-assigned-variable
        # left, right, width must added up to page_txt_width
        if self.prop_left is None or self.prop_right is None:
            return
        page_txt_width = self._prop_page_text_size.width
        tbl_min_width = 122
        left100 = self.prop_left.get_value_mm100()
        right100 = self.prop_right.get_value_mm100()

        width = page_txt_width - (left100 + right100)

        if width < tbl_min_width:  # min width:
            left = round((page_txt_width / 2) + (tbl_min_width / 2))
            right = left
            self._set(self._props.right, right)
            self._set(self._props.left, left)
            self._set(self._props.width, tbl_min_width)
            mLo.Lo.print(
                f"{self.__class__.__name__}. Unable to calculate proper table width. Result is less than min table width. Setting to full width."
            )
            return
        self._set(self._props.right, right100)
        self._set(self._props.left, left100)
        self._set(self._props.width, width)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblManualLeftRight: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblManualLeftRight: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblManualLeftRight:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblManualLeftRight: ``TblManualLeftRight`` Instance.
        """
        return _TblManualLeftRight._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblCenterWidth._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblManualCenter(_TblManualLeftRight):
    """
    Sets the table width manually.
    """

    # all properties of init are used in this class

    # region Overrides
    def _set_width_properties(self, obj: Any) -> None:
        # get the left, right and center values.
        # Calculate the difference between current width and
        # existing width. Apply the values
        if self.prop_width is None:
            return
        page_txt_width = self._prop_page_text_size.width
        tbl_min_width = 122
        width = self.prop_width.get_value_mm100()
        if width < tbl_min_width:
            mLo.Lo.print("Unable to change Table width. New Width is less then table min width.")
            return
        tbl_width = int(mProps.Props.get(obj, self._props.width))
        tbl_left = int(mProps.Props.get(obj, self._props.left))
        tbl_right = int(mProps.Props.get(obj, self._props.right))

        diff = tbl_width - width

        left = tbl_left + round(diff / 2)
        right = tbl_right + round(diff / 2)

        # make and minor adjustments need from rounding.
        while width + left + right > page_txt_width:
            if left >= 0:
                left -= 1
            else:
                left += 1
            if width + left + right == page_txt_width:
                break
            if right >= 0:
                right -= 1
            else:
                right += 1

        while width + left + right < page_txt_width:
            if left >= 0:
                left += 1
            else:
                left -= 1
            if width + left + right == page_txt_width:
                break
            if right >= 0:
                right += 1
            else:
                right -= 1

        self._set(self._props.left, left)
        self._set(self._props.right, right)
        self._set(self._props.width, width)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblManualCenter: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblManualCenter: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblManualCenter:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblManualCenter: ``TblManualCenter`` Instance.
        """
        return _TblManualCenter._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblAuto], obj: Any, **kwargs) -> _TTblAuto:
        return _TblManualLeftRight._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


# endregion Table size in MM units

# region Table size in percentage units


class _TblRelLeftByWidth(_SharedAuto):
    """
    Relative Table size. Set table right margin using width as a percentage value.
    """

    # region Init
    def __init__(
        self,
        *,
        width: TblRelUnit = 100,
        left: TblRelUnit = 0,
        right: TblRelUnit = 0,
        above: TblRelUnit | None = None,
        below: TblRelUnit | None = None,
    ) -> None:
        """
        Constructor

        Args:
            width (TblRelUnit, optional): Width value as a percentage from ``1`` and ``100``. Default ``100``.
            left (TblRelUnit, optional): Spacing Left value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            right (TblRelUnit, optional): Spacing Right value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            above (TblRelUnit, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (TblRelUnit, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        # right is omitted from constructor because it is (100 - width)
        # width and right are calculated and stored as 1/100th mm
        super().__init__(width=width, left=left, right=right, above=above, below=below)

    # endregion Init

    # region internal methods
    def _set_width_properties(self) -> None:
        page_width = self._prop_page_text_size.width
        self._set(self._props.left, 0)
        if self.prop_width.value == 100:
            self._set(self._props.width, page_width)
            self._set(self._props.right, 0)
            return
        width_factor = self.prop_width.value / 100
        width = round(page_width * width_factor)
        right = round(page_width - width)
        while right + width > page_width:
            # just in case rounding caused total to be more than page_width
            right = right - 1

        self._set(self._props.width, width)
        self._set(self._props.right, right)

    # endregion internal methods

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.LEFT.value)

    def _set_defaults(self) -> None:
        if not self._has(self._props.rel_width):
            self._set(self._props.rel_width, self.prop_width.value)
        super()._set_defaults()

    def apply(self, obj: Any, **kwargs) -> None:
        self._set_width_properties()
        return super().apply(obj, **kwargs)

    def _set_props_from_obj(self, obj: Any) -> None:
        super()._set_props_from_obj(obj)

        rel = self._get_relative_values()
        self._prop_width = rel.balance
        self._prop_left = rel.left
        self._prop_right = rel.right

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblRelLeftByWidth: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelLeftByWidth: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelLeftByWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelLeft: ``TblRelLeft`` Instance.
        """
        return _TblRelLeftByWidth._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblRelLeftByWidth], obj: Any, **kwargs) -> _TTblRelLeftByWidth:
        return _SharedAuto._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods

    # region properties
    @property
    def prop_width(self) -> Intensity:
        """
        Gets/Sets Width.
        """
        return Intensity(self._prop_width)

    @prop_width.setter
    def prop_width(self, value: int | Intensity):
        val = Intensity(int(value))
        self._prop_width = 1 if val.value == 0 else val.value

    @property
    def prop_left(self) -> Intensity:
        """
        Gets/Sets Left.
        """
        return Intensity(self._prop_left)

    @prop_left.setter
    def prop_left(self, value: int | Intensity):
        self._prop_left = Intensity(int(value)).value

    @property
    def prop_right(self) -> Intensity:
        """
        Gets/Sets Right.
        """
        return Intensity(self._prop_right)

    @prop_right.setter
    def prop_right(self, value: int | Intensity):
        self._prop_right = Intensity(int(value)).value

    # endregion properties


class _TblRelLeftByRight(_TblRelLeftByWidth):
    """
    Relative Table size. Set table right margin using right as a percentage value.
    """

    # region internal methods
    def _set_width_properties(self) -> None:
        page_txt_width = self._prop_page_text_size.width
        tbl_min_width100 = 122
        self._set(self._props.left, 0)
        if self.prop_right.value >= 99:
            self._set(self._props.rel_width, 1)
            self._set(self._props.width, tbl_min_width100)
            self._set(self._props.right, page_txt_width - tbl_min_width100)
            return
        right_factor = self.prop_right.value / 100
        right = round(page_txt_width * right_factor)
        width = page_txt_width - right
        while width + right > page_txt_width:
            # just in case rounding caused total to be more than page_width
            width = width - 1
        if width <= tbl_min_width100:
            self._set(self._props.rel_width, 1)
            self._set(self._props.width, tbl_min_width100)
            self._set(self._props.right, page_txt_width - tbl_min_width100)
            return
        # use ceil to make sure that width is at least 1
        width_per = math.ceil((width / page_txt_width) * 100)

        self._set(self._props.right, right)
        self._set(self._props.width, width)
        self._set(self._props.rel_width, width_per)

    # endregion internal methods

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.LEFT.value)

    def _set_defaults(self) -> None:
        if not self._has(self._props.rel_width):
            self._set(self._props.rel_width, self.prop_width.value)
        super()._set_defaults()

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblRelLeftByWidth: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelLeftByWidth: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelLeftByWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelLeft: ``TblRelLeft`` Instance.
        """
        return _TblRelLeftByRight._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblRelLeftByWidth], obj: Any, **kwargs) -> _TTblRelLeftByWidth:
        return _TblRelLeftByWidth._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblRelFromLeft(_TblRelLeftByWidth):
    """
    Relative Table size. Sets the table relative width from left using percentage values.
    """

    # width, left, above and below properties are used in this class
    # only width and left are used directly

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.FROM_LEFT.value)

    def _set_width_properties(self) -> None:
        # if left = 0 then right is right = round(page_width - width)
        page_txt_width = self._prop_page_text_size.width
        if self.prop_width.value == 100:
            self._set(self._props.width, page_txt_width)
            self._set(self._props.right, 0)
            self._set(self._props.left, 0)
            return

        if self.prop_left.value + self.prop_width.value > 100:
            # total of left and width must not be more then 100 percent.
            left_per = 100 - self.prop_width.value
        else:
            left_per = self.prop_left.value

        width_factor = self.prop_width.value / 100
        left_factor = 0 if left_per == 0 else left_per / 100
        width = round(page_txt_width * width_factor)
        left = 0 if left_factor == 0 else round(page_txt_width * left_factor)
        while left + width > page_txt_width:
            # just in case rounding caused total to be more than page_width
            left -= 1
        right = page_txt_width - (width + left)

        self._set(self._props.width, width)
        self._set(self._props.left, left)
        self._set(self._props.right, right)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblRelFromLeft: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelFromLeft: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelFromLeft:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelFromLeft: ``TblRelFromLeft`` Instance.
        """
        return _TblRelFromLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblRelLeftByWidth], obj: Any, **kwargs) -> _TTblRelLeftByWidth:
        return _TblRelLeftByWidth._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblRelRightByWidth(_TblRelLeftByWidth):
    """
    Relative Table size. Set table left margin using width as a percentage value.
    """

    # width, left, above and below properties are used in this class
    # only width and left are used directly

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.RIGHT.value)

    def _set_width_properties(self) -> None:
        # sourcery skip: extract-duplicate-method
        page_txt_width = self._prop_page_text_size.width
        if self.prop_width.value == 100:
            self._set(self._props.width, page_txt_width)
            self._set(self._props.left, 0)
            self._set(self._props.right, 0)
            return
        width_factor = self.prop_width.value / 100
        width = round(page_txt_width * width_factor)
        left = round(page_txt_width - width)
        while left + width > page_txt_width:
            # just in case rounding caused total to be more than page_width
            left = left - 1

        right = page_txt_width - width - left
        self._set(self._props.width, width)
        self._set(self._props.left, left)
        self._set(self._props.right, right)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblRelRightByWidth: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelRightByWidth: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelRightByWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelRight: ``TblRelRight`` Instance.
        """
        return _TblRelRightByWidth._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblRelLeftByWidth], obj: Any, **kwargs) -> _TTblRelLeftByWidth:
        return _TblRelLeftByWidth._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblRelRightByLeft(_TblRelRightByWidth):
    """
    Relative Table size. Set table left margin using left as a percentage value.
    """

    # region internal methods
    def _set_width_properties(self) -> None:
        page_txt_width = self._prop_page_text_size.width
        self._set(self._props.right, 0)
        if self.prop_left.value >= 99:
            self._set(self._props.rel_width, 1)
            tbl_min_width100 = 122
            self._set(self._props.width, tbl_min_width100)
            self._set(self._props.left, page_txt_width - tbl_min_width100)
            return
        left_factor = self.prop_left.value / 100
        left = round(page_txt_width * left_factor)
        width = page_txt_width - left
        while width + left > page_txt_width:
            # just in case rounding caused total to be more than page_width
            width -= 1

        # use ceil to make sure that width is at least 1
        width_per = math.ceil((width / page_txt_width) * 100)

        self._set(self._props.left, left)
        self._set(self._props.width, width)
        self._set(self._props.rel_width, width_per)

    # endregion internal methods

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.RIGHT.value)

    def _set_defaults(self) -> None:
        if not self._has(self._props.rel_width):
            self._set(self._props.rel_width, self.prop_width.value)
        super()._set_defaults()

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblRelLeftByWidth: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelLeftByWidth: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelLeftByWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelLeft: ``TblRelLeft`` Instance.
        """
        return _TblRelRightByLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblRelLeftByWidth], obj: Any, **kwargs) -> _TTblRelLeftByWidth:
        return _TblRelRightByWidth._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


class _TblRelCenter(_TblRelLeftByWidth):
    """
    Relative Table size. Set table left adn right margins using width as a percentage value.
    """

    # width, left, above and below properties are used in this class
    # only width and left are used directly

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, TableAlignKind.CENTER.value)

    def _set_width_properties(self) -> None:
        # sourcery skip: extract-duplicate-method
        page_width = self._prop_page_text_size.width
        if self.prop_width.value == 100:
            self._set(self._props.width, page_width)
            self._set(self._props.left, 0)
            self._set(self._props.right, 0)
            return

        width_factor = self.prop_width.value / 100
        width = round(page_width * width_factor)
        left_right = round(page_width - width)

        if left_right == 0:
            left = 0
            right = 0
        else:
            left = round(left_right / 2)
            right = left

        while right + left + width > page_width:
            # just in case rounding caused total to be more than page_width
            width = width - 1

        self._set(self._props.width, width)
        self._set(self._props.left, left)
        self._set(self._props.right, right)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> _TblRelCenter: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelCenter: ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> _TblRelCenter:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelCenter: ``TblRelCenter`` Instance.
        """
        return _TblRelCenter._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(clazz: Type[_TTblRelLeftByWidth], obj: Any, **kwargs) -> _TTblRelLeftByWidth:
        return _TblRelLeftByWidth._from_obj(clazz, obj, **kwargs)

    # endregion from_obj()
    # endregion static methods


# endregion Table size in percentage units


# region TableProperties Class
class TableProperties(StyleMulti):
    """
    Table options.

    Table horizontal position can be modified in many ways.
    This class generally follow how value are entered into the Table Properties dialog box of writer.

    ``name``, ``above`` and ``below`` parameters values are accepted for all states.

    When ``relative`` is ``True`` the following are required.

    .. cssclass:: ul-list

        - ``align=TableAlignKind.AUTO``, Not supported. Raises ``ValueError``.
        - ``align=TableAlignKind.CENTER``, ``width`` and ``left`` are required.
        - ``align=TableAlignKind.FROM_LEFT``, ``width`` and ``left`` are required.
        - ``align=TableAlignKind.LEFT``, ``width`` or ``right`` are required.
        - ``align=TableAlignKind.MANUAL``, Not supported. Raises ``ValueError``.
        - ``align=TableAlignKind.RIGHT``, ``width`` or ``left`` are required.

    When ``relative`` is ``False`` the following are required.

    .. cssclass:: ul-list

        - ``align=TableAlignKind.AUTO``, no extra parameters are required.
        - ``align=TableAlignKind.CENTER``, ``width`` or ``left`` are required.
        - ``align=TableAlignKind.FROM_LEFT``, ``width`` or ``left`` are required.
        - ``align=TableAlignKind.LEFT``, ``width`` or ``right`` is required.
        - ``align=TableAlignKind.MANUAL``, ``width``, ``left`` and ``right`` are required.
        - ``align=TableAlignKind.RIGHT``, ``width`` or ``left`` is required.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        name: str | None = None,
        width: TblAbsUnit | TblRelUnit | None = None,
        left: TblAbsUnit | TblRelUnit | None = None,
        right: TblAbsUnit | TblRelUnit | None = None,
        above: TblAbsUnit | TblRelUnit | None = None,
        below: TblAbsUnit | TblRelUnit | None = None,
        align: TableAlignKind | None = None,
        relative: bool = False,
    ) -> None:
        """
        Constructor

        Args:
            name (str, optional): Specifies frame name. Space are NOT allowed in names
            width (TblAbsUnit, TblRelUnit, optional): Specifies table Width.
            left (TblAbsUnit, TblRelUnit, optional): Specifies table Left.
            right (TblAbsUnit, TblRelUnit, optional): Specifies table Right.
            above (TblAbsUnit, TblRelUnit, optional): Specifies table spacing above.
            below (TblAbsUnit, TblRelUnit, optional): Specifies table spacing below.
            align (TableAlignKind, optional): Specifies table alignment.
            relative (bool, optional): Specifies if table horizontal values are in percentages or ``mm`` units.

        Raises:
            ValueError: If name contains spaces.

        Returns:
            None:
        """

        super().__init__()
        if name:
            name = name.strip()
            if " " in name:
                raise ValueError("Name is not allow to contains space characters.")
            self.prop_name = name

        self._prop_align = align
        self._prop_relative = relative

        if align is not None:
            size_obj = self._get_size_class(
                relative=relative,
                align=align,
                width=width,
                left=left,
                right=right,
                above=above,
                below=below,
            )
            self._set_style("size", size_obj)

    # endregion Init

    # region Overrides

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)
        # for some reason setting Name property raises "UnknownPropertyException" when "setPropertyValue()" is used (Which Props.set() uses).
        # However, setting Name via setattr() works fine.
        # for this reason this class cancels setting of Name property and sets it via setattr() here.
        if self._has(self._props.name) and hasattr(obj, self._props.name):
            name = getattr(obj, self._props.name)
            if name != self.prop_name:
                # not case sensitive
                setattr(obj, self._props.name, self.prop_name)

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == self._props.name:
            event_args.cancel = True
            event_args.handled = True
            # see bug specified in apply() method.
        super().on_property_setting(source, event_args)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = _get_default_tbl_services()
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region internal methods
    def _get_size_class(
        self,
        *,
        relative: bool,
        align: TableAlignKind,
        width: TblAbsUnit | TblRelUnit | None = None,
        left: TblAbsUnit | TblRelUnit | None = None,
        right: TblAbsUnit | TblRelUnit | None = None,
        above: TblAbsUnit | TblRelUnit | None = None,
        below: TblAbsUnit | TblRelUnit | None = None,
    ) -> _TblAuto:
        if relative:
            return self._get_size_rel_class(
                align=align,
                width=cast(Union[TblRelUnit, None], width),
                left=cast(Union[TblRelUnit, None], left),
                right=cast(Union[TblRelUnit, None], right),
                above=cast(Union[TblRelUnit, None], above),
                below=cast(Union[TblRelUnit, None], below),
            )
        return self._get_size_abs_class(
            align=align,
            width=cast(Union[TblAbsUnit, None], width),
            left=cast(Union[TblAbsUnit, None], left),
            right=cast(Union[TblAbsUnit, None], right),
            above=cast(Union[TblAbsUnit, None], above),
            below=cast(Union[TblAbsUnit, None], below),
        )

    def _get_size_abs_class(
        self,
        *,
        align: TableAlignKind,
        width: TblAbsUnit | None = None,
        left: TblAbsUnit | None = None,
        right: TblAbsUnit | None = None,
        above: TblAbsUnit | None = None,
        below: TblAbsUnit | None = None,
    ) -> _TblAuto:
        def check_req(*args: Any) -> bool:
            for arg in args:
                if arg is None:
                    return False
            return True

        if align == TableAlignKind.AUTO:
            return _TblAuto(above=above, below=below, _cattribs=self._get_tbl_cattribs())  # type: ignore
        if align == TableAlignKind.CENTER:
            if width is None:
                ck = check_req(left)
                if not ck:
                    raise ValueError(
                        f"left or width are required when align is set to {align.name} and relative value is False."
                    )
                return _TblCenterLeft(above=above, below=below, left=left, _cattribs=self._get_tbl_cattribs())  # type: ignore
            else:
                return _TblCenterWidth(above=above, below=below, width=width, _cattribs=self._get_tbl_cattribs())  # type: ignore
        if align == TableAlignKind.FROM_LEFT:
            if width is None:
                ck = check_req(left)
                if not ck:
                    raise ValueError(
                        f"left or width are required when align is set to {align.name} and relative value is False."
                    )
                return _TblFromLeft(above=above, below=below, left=left, _cattribs=self._get_tbl_cattribs())  # type: ignore
            else:
                return _TblFromLeftWidth(above=above, below=below, width=width, _cattribs=self._get_tbl_cattribs())  # type: ignore
        if align == TableAlignKind.LEFT:
            if width is None:
                ck = check_req(right)
                if not ck:
                    raise ValueError(
                        f"right or width are required when align is set to {align.name} and relative value is False."
                    )
                return _TblLeft(above=above, below=below, right=right, _cattribs=self._get_tbl_cattribs())  # type: ignore
            else:
                return _TblLeftWidth(above=above, below=below, width=width, _cattribs=self._get_tbl_cattribs())  # type: ignore
        if align == TableAlignKind.RIGHT:
            if width is None:
                ck = check_req(left)
                if not ck:
                    raise ValueError(
                        f"left or width are required when align is set to {align.name} and relative value is False."
                    )
                return _TblRight(above=above, below=below, left=left, _cattribs=self._get_tbl_cattribs())  # type: ignore
            else:
                return _TblRightWidth(above=above, below=below, width=width, _cattribs=self._get_tbl_cattribs())  # type: ignore
        if align == TableAlignKind.MANUAL:
            if width is None:
                ck = check_req(left, right)
                if not ck:
                    raise ValueError(
                        f"left and right or width are required when align is set to {align.name} and relative value is False."
                    )
                return _TblManualLeftRight(
                    above=above, below=below, left=left, right=right, _cattribs=self._get_tbl_cattribs()  # type: ignore
                )
            else:
                return _TblManualCenter(above=above, below=below, width=width, _cattribs=self._get_tbl_cattribs())  # type: ignore
        raise ValueError("Align Value is Unknown")

    def _get_size_rel_class(
        self,
        *,
        align: TableAlignKind,
        width: TblRelUnit | None = None,
        left: TblRelUnit | None = None,
        right: TblRelUnit | None = None,
        above: TblRelUnit | None = None,
        below: TblRelUnit | None = None,
    ) -> _TTblAuto | _TTblRelLeftByWidth:  # type: ignore
        def check_req(*args: Any) -> bool:
            for arg in args:
                if arg is None:
                    return False
            return True

        if align == TableAlignKind.CENTER:
            if width is None:
                ck = check_req(left)
                if not ck:
                    raise ValueError(
                        f"left or width are required when align is set to {align.name} and relative value is True."
                    )
                return cast(
                    _TTblRelLeftByWidth,
                    _TblRelCenter(above=above, below=below, left=left, _cattribs=self._get_tbl_cattribs()),  # type: ignore
                )
            else:
                return cast(
                    _TTblAuto,
                    _TblManualCenter(above=above, below=below, width=width, _cattribs=self._get_tbl_cattribs()),  # type: ignore
                )
        if align == TableAlignKind.FROM_LEFT:
            ck = check_req(width, left)
            if not ck:
                raise ValueError(
                    f"width and left are required when align is set to {align.name} and relative value is True."
                )
            return cast(
                _TTblRelLeftByWidth,
                _TblRelFromLeft(above=above, below=below, width=width, left=left, _cattribs=self._get_tbl_cattribs()),  # type: ignore
            )
        if align == TableAlignKind.LEFT:
            if width is None:
                ck = check_req(right)
                if not ck:
                    raise ValueError(
                        f"right or width are required when align is set to {align.name} and relative value is True."
                    )
                return cast(
                    _TTblRelLeftByWidth,
                    _TblRelLeftByRight(above=above, below=below, right=right, _cattribs=self._get_tbl_cattribs()),  # type: ignore
                )
            else:
                return cast(
                    _TTblRelLeftByWidth,
                    _TblRelLeftByWidth(above=above, below=below, width=width, _cattribs=self._get_tbl_cattribs()),  # type: ignore
                )
        if align == TableAlignKind.RIGHT:
            if width is None:
                ck = check_req(left)
                if not ck:
                    raise ValueError(
                        f"left or width are required when align is set to {align.name} and relative value is True."
                    )
                return cast(
                    _TTblRelLeftByWidth,
                    _TblRelRightByLeft(above=above, below=below, left=left, _cattribs=self._get_tbl_cattribs()),  # type: ignore
                )
            else:
                return cast(
                    _TTblRelLeftByWidth,
                    _TblRelRightByWidth(above=above, below=below, width=width, _cattribs=self._get_tbl_cattribs()),  # type: ignore
                )
        if align == TableAlignKind.AUTO:
            raise ValueError('align must not be set to "TableAlignKind.AUTO" when relative is set to False')
        if align == TableAlignKind.MANUAL:
            raise ValueError('align must not be set to "TableAlignKind.MANUAL" when relative is set to False')
        raise ValueError("Align Value is Unknown")

    def _get_tbl_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
            "_props_internal_attributes": self._props,
        }

    # endregion internal methods

    # region static methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTableProperties], obj: Any) -> _TTableProperties: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTableProperties], obj: Any, **kwargs) -> _TTableProperties: ...

    @classmethod
    def from_obj(cls: Type[_TTableProperties], obj: Any, **kwargs) -> _TTableProperties:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TableProperties: Instance that represents Table options.
        """
        # sourcery skip: merge-else-if-into-elif
        # this nu is only used to get Property Name

        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        try:
            inst._set(inst._props.name, mProps.Props.get(obj, inst._props.name))
        except mEx.PropertyNotFoundError:
            # there is a bug. See apply()
            if hasattr(obj, inst._props.name):
                inst._set(inst._props.name, getattr(obj, inst._props.name))
            else:
                raise
        hori = int(mProps.Props.get(obj, inst._props.hori_orient))
        rel = bool(mProps.Props.get(obj, inst._props.is_rel))
        inst._prop_relative = rel
        inst._prop_align = TableAlignKind(hori)

        if rel:
            if hori == HoriOrientation.LEFT:
                size_obj = _TblRelLeftByWidth.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.LEFT_AND_WIDTH:
                size_obj = _TblRelFromLeft.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.RIGHT:
                size_obj = _TblRelRightByWidth.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            else:
                size_obj = _TblRelCenter.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
        else:
            if hori == HoriOrientation.FULL:
                size_obj = _TblAuto.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.CENTER:
                size_obj = _TblCenterWidth.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.LEFT:
                size_obj = _TblLeft.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.RIGHT:
                size_obj = _TblRight.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.LEFT_AND_WIDTH:
                size_obj = _TblFromLeft.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            else:
                size_obj = _TblManualLeftRight.from_obj(obj, _cattribs=inst._get_tbl_cattribs())

        inst._set_style("size", size_obj)

        # prev, next not currently working
        return inst

    # endregion from_obj()

    # endregion static methods

    # region Methods
    def get_width_mm(self) -> UnitMM | None:
        """
        Gets Width in ``mm`` units.

        When class is constructed using relative values this method will still
        return ``mm`` units.

        Returns:
            UnitMM | None:

        Note:
            This method may return None if ``apply()`` has not yet been called.
        """
        po = cast(_TSharedAuto, self._prop_obj)  # type: ignore
        if po is None:
            return None
        pv = cast(int, po._get(self._props.width))
        return None if pv is None else UnitMM.from_mm100(pv)

    def get_left_mm(self) -> UnitMM | None:
        """
        Gets left in ``mm`` units.

        When class is constructed using relative values this method will still
        return ``mm`` units.

        Returns:
            UnitMM | None:

        Note:
            This method may return None if ``apply()`` has not yet been called.
        """
        po = cast(_TSharedAuto, self._prop_obj)  # type: ignore
        if po is None:
            return None
        pv = cast(int, po._get(self._props.left))
        return None if pv is None else UnitMM.from_mm100(pv)

    def get_right_mm(self) -> UnitMM | None:
        """
        Gets right in ``mm`` units.

        When class is constructed using relative values this method will still
        return ``mm`` units.

        Returns:
            UnitMM | None:

        Note:
            This method may return None if ``apply()`` has not yet been called.
        """
        po = cast(_TSharedAuto, self._prop_obj)  # type: ignore
        if po is None:
            return None
        pv = cast(int, po._get(self._props.right))
        return None if pv is None else UnitMM.from_mm100(pv)

    # endregion Methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.TABLE
        return self._format_kind_prop

    @property
    def prop_name(self) -> str | None:
        """Gets/Sets name"""
        return self._get(self._props.name)

    @prop_name.setter
    def prop_name(self, value: str | None) -> None:
        if value is None:
            self._remove(self._props.name)
            return
        self._set(self._props.name, value)

    @property
    def _prop_obj(self) -> Any:
        """Gets inner instance if exist."""
        try:
            return self._prop_inner_obj
        except AttributeError:
            val = self._get_style("size")
            self._prop_inner_obj = None if val is None else val.style
        return self._prop_inner_obj

    @property
    def prop_align(self) -> TableAlignKind | None:
        """Gets Alignment value"""
        return self._prop_align

    @property
    def prop_relative(self) -> bool:
        """Gets relative value"""
        return self._prop_relative

    @property
    def prop_width(self) -> UnitMM | Intensity | None:
        """
        Gets width value.

        Returns:
            UnitMM | Intensity | None: When ``relative`` is ``True`` ``Intensity`` or ``None``; Otherwise, ``UnitMM`` or None.

        See Also:
            :py:meth:`~.table_properties.TableProperties.get_width_mm`.
        """
        po = cast(_TSharedAuto, self._prop_obj)  # type: ignore
        return None if po is None else po.prop_width

    @property
    def prop_left(self) -> UnitMM | Intensity | None:
        """
        Gets left value.

        Returns:
            UnitMM | Intensity | None: When ``relative`` is ``True`` ``Intensity`` or ``None``; Otherwise, ``UnitMM`` or None.

        See Also:
            :py:meth:`~.table_properties.TableProperties.get_left_mm`.
        """
        po = cast(_TSharedAuto, self._prop_obj)  # type: ignore
        return None if po is None else po.prop_left

    @property
    def prop_right(self) -> UnitMM | Intensity | None:
        """
        Gets right value.

        Returns:
            UnitMM | Intensity | None: When ``relative`` is ``True`` ``Intensity`` or ``None``; Otherwise, ``UnitMM`` or None.

        See Also:
            :py:meth:`~.table_properties.TableProperties.get_right_mm`.
        """
        po = cast(_TSharedAuto, self._prop_obj)  # type: ignore
        return None if po is None else po.prop_right

    @property
    def prop_above(self) -> UnitMM | None:
        """Gets above value"""
        po = cast(_TSharedAuto, self._prop_obj)  # type: ignore
        if po is None:
            return None
        pv = cast(int, po._get(self._props.top))
        return None if pv is None else UnitMM.from_mm100(pv)

    @property
    def prop_below(self) -> UnitMM | None:
        """Gets below value"""
        po = cast(_TSharedAuto, self._prop_obj)  # type: ignore
        if po is None:
            return None
        pv = cast(int, po._get(self._props.bottom))
        return None if pv is None else UnitMM.from_mm100(pv)

    @property
    def _props(self) -> TablePropertiesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _get_default_tbl_props()
        return self._props_internal_attributes

    # endregion properties


# endregion TableProperties Class

__all__ = ("TblAbsUnit", "TblRelUnit", "TableAlignKind", "TableProperties")
