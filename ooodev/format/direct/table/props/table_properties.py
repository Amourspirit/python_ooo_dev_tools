from __future__ import annotations
from typing import cast, overload
from typing import Any, Tuple, Type, TypeVar, NamedTuple
import uno
from ooo.dyn.text.hori_orientation import HoriOrientation

from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....exceptions import ex as mEx
from .....proto.style_obj import StyleObj
from .....proto.unit_obj import UnitObj
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.data_type.intensity import Intensity
from .....utils.data_type.size import Size
from .....utils.data_type.unit_mm import UnitMM
from ....direct.common.abstract.abstract_document import AbstractDocument
from ....kind.format_kind import FormatKind
from ....style_base import StyleMulti
from ...common.props.table_properties_props import TablePropertiesProps
from .....utils.unit_convert import UnitConvert

_TTableProperties = TypeVar(name="_TTableProperties", bound="TableProperties")
_TTblAuto = TypeVar(name="_TTblAuto", bound="TblAuto")


class RelVals(NamedTuple):
    """Relative values. Total expected to add up to ``100``"""

    left: int
    right: int
    balance: int


def _get_default_tbl_props() -> TablePropertiesProps:
    return TablePropertiesProps(
        name="Name",
        width="Width",
        left="LeftMargin",
        top="TopMargin",
        right="RightMargin",
        bottom="BottomMargin",
        is_rel="IsWidthRelative",
        hori_orient="HoriOrient",
    )


def _get_default_tbl_services() -> Tuple[str]:
    return ("com.sun.star.text.TextTable",)


class TblPropObj(StyleObj):
    """
    Protcol class for Table Properties

    Supported Classes:

    .. cssclass:: ul-list

        - :py:class:`~.format.direct.table.props.table_properties.TblAuto`
        - :py:class:`~.format.direct.table.props.table_properties.TblWidth`
        - :py:class:`~.format.direct.table.props.table_properties.TblLeft`
        - :py:class:`~.format.direct.table.props.table_properties.TblRight`
        - :py:class:`~.format.direct.table.props.table_properties.TblCenter`
        - :py:class:`~.format.direct.table.props.table_properties.TblFromLeft`
        - :py:class:`~.format.direct.table.props.table_properties.TblFromLeftWidth`
        - :py:class:`~.format.direct.table.props.table_properties.TblManual`
        - :py:class:`~.format.direct.table.props.table_properties.TblRelLeft`
        - :py:class:`~.format.direct.table.props.table_properties.TblRelRight`
        - :py:class:`~.format.direct.table.props.table_properties.TblRelCenter`

    """

    pass


# region Table size in MM units


class TblAuto(AbstractDocument):
    """
    Automatically set table width.
    """

    # region Init
    def __init__(self, above: float | UnitObj = 0, below: float | UnitObj = 0) -> None:
        """
        Constructor

        Args:
            above (float, UnitObj, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (float, UnitObj, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        super().__init__()
        self.prop_above = above
        self.prop_below = below
        self._post_init()

    # endregion Init

    # region internal methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.FULL)

    def _get_relative_values(self) -> RelVals:
        left100: int = self._get(self._props.left)
        right100: int = self._get(self._props.right)
        width100: int = self._get(self._props.width)
        total100 = left100 + right100 + width100
        if left100 == 0:
            left = 0
        else:
            left = round(total100 / left100)
        if right100 == 0:
            right = 0
        else:
            right = round(total100 / right100)
        width = 100 - (left + right)
        return RelVals(left=left, right=right, balance=width)

    # endregion internal methods

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = _get_default_tbl_services()
        return self._supported_services_values

    def _set_defaults(self) -> None:
        if not self._has(self._props.width):
            self._set(self._props.width, self._prop_page_text_size.width)
        if not self._has(self._props.left):
            self._set(self._props.left, 0)
        if not self._has(self._props.right):
            self._set(self._props.right, 0)
        if not self._has(self._props.is_rel):
            self._set(self._props.is_rel, False)

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies Styling to object

        Args:
            obj (object): UNO Table Object

        Returns:
            None:
        """
        self._set_defaults()
        return super().apply(obj, **kwargs)

    # endregion overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblAuto:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblAuto:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblAuto:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblAuto: ``TblAuto`` Instance.
        """
        return TblAuto._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Options: Instance that represents Frame Wrap Settings.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        inst._set(inst._props.left, mProps.Props.get(obj, inst._props.left))
        inst._set(inst._props.right, mProps.Props.get(obj, inst._props.right))
        inst._set(inst._props.top, mProps.Props.get(obj, inst._props.top))
        inst._set(inst._props.bottom, mProps.Props.get(obj, inst._props.bottom))
        inst._set(inst._props.width, mProps.Props.get(obj, inst._props.width))
        inst._set(inst._props.hori_orient, mProps.Props.get(obj, inst._props.hori_orient))
        inst._set(inst._props.is_rel, mProps.Props.get(obj, inst._props.is_rel))
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
    def prop_above(self) -> UnitMM:
        """
        Gets/Sets above.
        """
        return UnitMM.from_mm100(self._get(self._props.top))

    @prop_above.setter
    def prop_above(self, value: float | UnitObj):
        try:
            val = value.get_value_mm100()
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)
        self._set(self._props.top, val)

    @property
    def prop_below(self) -> UnitMM:
        """
        Gets/Sets below.
        """
        return UnitMM.from_mm100(self._get(self._props.bottom))

    @prop_below.setter
    def prop_below(self, value: float | UnitObj):
        try:
            val = value.get_value_mm100()
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)
        self._set(self._props.bottom, val)

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


class TblWidth(TblAuto):
    """
    Sets the table width using a width value and centers the table.
    """

    # region Init
    def __init__(
        self,
        width: float | UnitObj = 0,
        above: float | UnitObj = 0,
        below: float | UnitObj = 0,
    ) -> None:
        """
        _summary_

        Args:
            width (float, UnitObj): Table Width value in ``mm`` units or :ref:`proto_unit_obj`.
            above (float, UnitObj, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (float, UnitObj, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        super().__init__(above=above, below=below)
        self.prop_width = width

    # endregion Init

    # region internal Methods
    def _set_width_properties(self, obj: object) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
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
        self._set(self._props.hori_orient, HoriOrientation.CENTER)

    def copy(self: _TTblAuto, **kwargs) -> _TTblAuto:
        """Gets a copy of the instance"""
        cp = super().copy(**kwargs)
        cp._prop_width = self._prop_width
        return cp

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies Styling to object

        Args:
            obj (object): UNO Table Object

        Returns:
            None:
        """
        self._set_width_properties(obj)
        return super().apply(obj, **kwargs)

    # endregion overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblWidth:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblWidth:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblWidth: ``TblWidth`` Instance.
        """
        return TblWidth._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblAuto._from_obj(cls, obj, **kwargs)
        inst._prop_width = int(inst._get(inst._props.width))
        return inst

    # endregion from_obj()
    # endregion static methods

    # region properties

    @property
    def prop_width(self) -> UnitMM:
        """
        Gets/Sets Width.
        """
        return UnitMM.from_mm100(self._prop_width)

    @prop_width.setter
    def prop_width(self, value: float | UnitObj):
        try:
            val = value.get_value_mm100()
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)
        # min table width seems to be in the area of 1.22 mm
        if val < 122:
            val = 122
        self._prop_width = val

    # endregion properties


class TblLeft(TblAuto):
    """
    Sets the right table margin.
    """

    # region Init
    def __init__(
        self,
        margin: float | UnitObj = 0,
        above: float | UnitObj = 0,
        below: float | UnitObj = 0,
    ) -> None:
        """
        Constructor

        Args:
            margin (float, UnitObj, optional): Margin Value in ``mm`` units or :ref:`proto_unit_obj`.
            above (float, UnitObj, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (float, UnitObj, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        super().__init__(above=above, below=below)
        self.prop_margin = margin

    # endregion Init

    # region internal Methods

    def _set_width_properties(self, obj: object) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
        self._set(self._props.left, 0)
        page_txt_width = self._prop_page_text_size.width
        margin = self.prop_margin.get_value_mm100()

        if margin == 0:
            self._set(self._props.right, 0)
            self._set(self._props.width, page_txt_width)
            return

        # minimum width allowed for table seems to be 1.22 mm
        min_width100 = 122
        if margin > 0:
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
        self._set(self._props.hori_orient, HoriOrientation.LEFT)

    def copy(self: _TTblAuto, **kwargs) -> _TTblAuto:
        """Gets a copy of the instance"""
        cp = super().copy(**kwargs)
        cp._prop_margin = self._prop_margin
        return cp

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies Styling to object

        Args:
            obj (object): UNO Table Object

        Returns:
            None:
        """
        self._set_width_properties(obj)
        return super().apply(obj, **kwargs)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblLeft:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblLeft:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblLeft:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblLeft: ``TblLeft`` Instance.
        """

        return TblLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls, obj: object, **kwargs) -> TblLeft:
        inst = TblAuto._from_obj(cls, obj, **kwargs)
        inst._prop_margin = int(inst._get(inst._props.right))
        return inst

    # endregion from_obj()
    # endregion static methods

    # region properties

    @property
    def prop_margin(self) -> UnitMM:
        """
        Gets/Sets right.
        """
        return UnitMM.from_mm100(self._prop_margin)

    @prop_margin.setter
    def prop_margin(self, value: float | UnitObj):
        try:
            val = value.get_value_mm100()
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)
        self._prop_margin = val

    # endregion properties


class TblRight(TblLeft):
    """
    Sets the left table margin.
    """

    # region override Methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.RIGHT)

    def _set_width_properties(self, obj: object) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
        self._set(self._props.right, 0)
        page_txt_width = self._prop_page_text_size.width
        margin = self.prop_margin.get_value_mm100()

        if margin == 0:
            self._set(self._props.left, 0)
            self._set(self._props.width, page_txt_width)
            return

        # minimum width allowed for table seems to be 1.22 mm
        min_width100 = 122
        if margin > 0:
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
    def from_obj(cls, obj: object) -> TblRight:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRight:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRight:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRight: ``TblRight`` Instance.
        """
        return TblRight._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblLeft._from_obj(cls, obj, **kwargs)
        inst._prop_margin = int(inst._get(inst._props.left))
        return inst

    # endregion from_obj()
    # endregion static methods


class TblCenter(TblLeft):
    """
    Sets the table width using a margin value and centers the table.
    """

    # region override Methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.CENTER)

    def _set_width_properties(self, obj: object) -> None:
        # max right value is the neg value of page text area
        # max width is page text area * 2
        page_txt_width = self._prop_page_text_size.width
        margin = self.prop_margin.get_value_mm100()

        if margin == 0:
            self._set(self._props.left, 0)
            self._set(self._props.right, 0)
            self._set(self._props.width, page_txt_width)
            return

        # minimum width allowed for table seems to be 1.22 mm
        min_width100 = 122
        if margin > 0:
            if margin > page_txt_width + min_width100:
                margin = page_txt_width - min_width100
            width100 = page_txt_width - margin
            left_margin = round(margin / 2)
            right_margin = left_margin
            while left_margin + right_margin > margin:
                # just in case round added a digit or two
                right_margin -= 1

            self._set(self._props.width, width100)
            self._set(self._props.left, left_margin)
            self._set(self._props.right, right_margin)
            return

        # margin is negative

        if abs(margin) > page_txt_width:
            margin = -page_txt_width

        left_margin = round(margin / 2)
        right_margin = left_margin

        while left_margin + right_margin < margin:
            # just in case round added a digit or two
            right_margin += 1

        width100 = abs(margin) + page_txt_width
        self._set(self._props.width, width100)
        self._set(self._props.left, left_margin)
        self._set(self._props.right, right_margin)

    # endregion override Methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblRight:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRight:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRight:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRight: ``TblRight`` Instance.
        """
        return TblCenter._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblLeft._from_obj(cls, obj, **kwargs)
        return inst

    # endregion from_obj()
    # endregion static methods


class TblFromLeft(TblLeft):
    """
    Sets the table width from left using a margin value.
    """

    # region override Methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.LEFT_AND_WIDTH)

    def _set_width_properties(self, obj: object) -> None:
        # the total of width + left + right = page_txt_width
        # when left changes the width remains the same and right can become a neg value
        # if width exceeds page_txt_width then is set to page_txt_width, left and right become 0
        # min table width 1.22 mm

        # if left exceeds page_txt_width then left is set to page_txt_width, width is unchanged and right is set to neg width

        page_txt_width = self._prop_page_text_size.width
        margin = self.prop_margin.get_value_mm100()
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
    def from_obj(cls, obj: object) -> TblFromLeft:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblFromLeft:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblFromLeft:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblFromLeft: ``TblFromLeft`` Instance.
        """
        return TblFromLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblLeft._from_obj(cls, obj, **kwargs)
        inst._prop_margin = int(inst._get(inst._props.left))
        return inst

    # endregion from_obj()
    # endregion static methods


class TblFromLeftWidth(TblWidth):
    """
    Sets the table width from left using a width value.
    """

    # region override Methods
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.LEFT_AND_WIDTH)

    def _set_width_properties(self, obj: object) -> None:
        # the total of width + left + right = page_txt_width
        # when width changes left stays the same and right gets adjusted conditionally.
        #   if right is 0 then width will reduce left
        # if width exceeds page_txt_width then is set to page_txt_width, left and right become 0
        # min table width 1.22 mm

        page_txt_width = self._prop_page_text_size.width

        width100 = self.prop_width.get_value_mm()
        if width100 >= page_txt_width:
            self._set(self._props.right, 0)
            self._set(self._props.left, 0)
            self._set(self._props.width, page_txt_width)
            return

        left = int(mProps.Props.get(obj, self._props.left))
        right = page_txt_width - width100 - left

        if right < 0:
            left = left + right
            right = 0

        while left + right + width100 > page_txt_width:
            width100 -= 1

        if width100 < 122:  # min width:
            self._set(self._props.right, 0)
            self._set(self._props.left, 0)
            self._set(self._props.width, page_txt_width)
            mLo.Lo.print(f"{self.__class__.__name__}. Unable to calculate proper table width. Setting to full width.")
            return

        self._set(self._props.right, right)
        self._set(self._props.left, left)
        self._set(self._props.width, width100)

    # endregion override Methods

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblFromLeftWidth:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblFromLeftWidth:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblFromLeftWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblFromLeftWidth: ``TblFromLeftWidth`` Instance.
        """
        return TblFromLeftWidth._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblWidth._from_obj(cls, obj, **kwargs)
        return inst

    # endregion from_obj()
    # endregion static methods


class TblManual(TblWidth):
    """
    Sets the table width manually.
    """

    def __init__(
        self,
        width: float | UnitObj = 0,
        left: float | UnitObj = 0,
        right: float | UnitObj = 0,
        above: float | UnitObj = 0,
        below: float | UnitObj = 0,
    ) -> None:
        """
        _summary_

        Args:
            width (float, UnitObj): Table Width value in ``mm`` units or :ref:`proto_unit_obj`.
            left (float, UnitObj, optional): Left margin in ``mm`` units or :ref:`proto_unit_obj`. Defaults to ``0``.
            right (float, UnitObj, optional): Rigth margin in ``mm`` units or :ref:`proto_unit_obj`. Defaults to ``0``.
            above (float, UnitObj, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (float, UnitObj, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        super().__init__(width=width, above=above, below=below)
        self.prop_right = right
        self.prop_left = left

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.NONE)

    def copy(self: _TTblAuto, **kwargs) -> _TTblAuto:
        """Gets a copy of the instance."""
        cp = super().copy(**kwargs)
        cp._prop_right = self._prop_right
        cp._prop_left = self._prop_left
        return cp

    def _set_width_properties(self, obj: object) -> None:
        # left, right, width must added up to page_txt_width
        page_txt_width = self._prop_page_text_size.width
        width100 = self.prop_width.get_value_mm100()
        left100 = self.prop_left.get_value_mm100()
        right100 = self.prop_right.get_value_mm100()

        while width100 + left100 + right100 < page_txt_width:
            right100 += 1

        while width100 + left100 + right100 > page_txt_width:
            width100 -= 1

        if width100 < 122:  # min width:
            self._set(self._props.right, 0)
            self._set(self._props.left, 0)
            self._set(self._props.width, page_txt_width)
            mLo.Lo.print(f"{self.__class__.__name__}. Unable to calculate proper table width. Setting to full width.")
            return

        self._set(self._props.right, right100)
        self._set(self._props.left, left100)
        self._set(self._props.width, width100)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblWidth:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblWidth:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblWidth:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblManual: ``TblManual`` Instance.
        """
        # this nu is only used to get Property Name

        return TblManual._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblWidth._from_obj(cls, obj, **kwargs)
        inst._prop_right = int(inst._get(inst._props.right))
        inst._prop_left = int(inst._get(inst._props.left))
        return inst

    # endregion from_obj()
    # endregion static methods

    # region properties

    @property
    def prop_right(self) -> UnitMM:
        """
        Gets/Sets right.
        """
        return UnitMM.from_mm100(self._prop_right)

    @prop_right.setter
    def prop_right(self, value: float | UnitObj):
        try:
            val = value.get_value_mm100()
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)
        self._prop_right = val

    @property
    def prop_left(self) -> UnitMM:
        """
        Gets/Sets left.
        """
        return UnitMM.from_mm100(self._prop_left)

    @prop_left.setter
    def prop_left(self, value: float | UnitObj):
        try:
            val = value.get_value_mm100()
        except AttributeError:
            val = UnitConvert.convert_mm_mm100(value)
        self._prop_left = val

    # endregion properties


# endregion Table size in MM units

# region Table size in percentage units


class TblRelLeft(TblAuto):
    """
    Relative Table size. Set table right margin using width as a percentage value.
    """

    def __init__(self, width: int | Intensity = 100, above: float | UnitObj = 0, below: float | UnitObj = 0) -> None:
        """
        Constructor

        Args:
            width (int, Intensity, optional): Width value as a percentage from ``1`` and ``100``. Default ``100``.
            above (float, UnitObj, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (float, UnitObj, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        # right is ommited from constructor because it is (100 - width)
        # width and right are calculated and stored as 1/100th mm
        super().__init__(above=above, below=below)
        self.prop_width = width

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
        self._set(self._props.hori_orient, HoriOrientation.LEFT)

    def copy(self: _TTblAuto, **kwargs) -> _TTblAuto:
        """Gets a copy of the instance"""
        cp = super().copy(**kwargs)
        cp._width = self._width
        return cp

    def _set_defaults(self) -> None:
        self._set(self._props.is_rel, True)
        super()._set_defaults()

    def apply(self, obj: object, **kwargs) -> None:
        self._set_width_properties()
        return super().apply(obj, **kwargs)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblRelLeft:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRelLeft:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRelLeft:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelLeft: ``TblRelLeft`` Instance.
        """
        return TblRelLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblAuto._from_obj(cls, obj, **kwargs)
        rel = inst._get_relative_values()
        inst.prop_width = rel.balance
        return inst

    # endregion from_obj()
    # endregion static methods

    # region properties
    @property
    def prop_width(self) -> Intensity:
        """
        Gets/Sets Width.
        """
        return Intensity(self._width)

    @prop_width.setter
    def prop_width(self, value: int | Intensity):
        val = Intensity(int(value))
        if val.value == 0:
            # min value is always 1
            self._width = 1
        else:
            self._width = val.value

    # endregion properties


class TblRelFromLeft(TblRelLeft):
    """
    Relative Table size. Sets the table relative width from left using percentage values.
    """

    # region Init
    def __init__(
        self,
        width: int | Intensity = 100,
        left: int | Intensity = 0,
        above: float | UnitObj = 0,
        below: float | UnitObj = 0,
    ) -> None:
        """
        Constructor

        Args:
            width (int, Intensity, optional): Width value as a percentage from ``1`` and ``100``. Default ``100``.
            left (int, Intensity, optional): Left value as a percentage from ``0`` and ``100``. Default ``0``.
            above (float, UnitObj, optional): Spacing Above value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
            below (float, UnitObj, optional): Spacing Below value in ``mm`` units or :ref:`proto_unit_obj`. Default ``0``.
        """
        super().__init__(width=width, above=above, below=below)
        self.prop_left = left

    # endregion Init

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.LEFT_AND_WIDTH)

    def copy(self: _TTblAuto, **kwargs) -> _TTblAuto:
        """Gets an copy of the instance."""
        cp = super().copy(**kwargs)
        cp._left = self._left
        return cp

    def _set_width_properties(self) -> None:
        # if left = 0 then right is right = round(page_width - width)
        page_width = self._prop_page_text_size.width
        if self.prop_width.value == 100:
            self._set(self._props.width, page_width)
            self._set(self._props.right, 0)
            self._set(self._props.left, 0)
            return

        if self.prop_left.value + self.prop_width.value > 100:
            # total of left and width must no be more then 100 percent.
            left_per = 100 - self.prop_width.value
        else:
            left_per = self.prop_left.value

        width_factor = self.width.value / 100
        if left_per == 0:
            left_factor = 0
        else:
            left_factor = left_per / 100

        width = round(page_width * width_factor)
        if left_factor == 0:
            left = 0
        else:
            left = round(page_width * left_factor)
        while left + width > page_width:
            # just in case rounding caused total to be more than page_width
            left = left - 1

        right = page_width - (width + left)
        if right < 0:
            right = 0

        self._set(self._props.width, width)
        self._set(self._props.left, left)
        self._set(self._props.right, right)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblRelFromLeft:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRelFromLeft:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRelFromLeft:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelFromLeft: ``TblRelFromLeft`` Instance.
        """
        return TblRelFromLeft._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblRelLeft._from_obj(cls, obj, **kwargs)
        rel = inst._get_relative_values()
        inst.prop_left = rel.left
        return inst

    # endregion from_obj()
    # endregion static methods

    # region properties
    @property
    def prop_left(self) -> Intensity:
        """
        Gets/Sets Left.
        """
        return Intensity(self._left)

    @prop_left.setter
    def prop_left(self, value: int | Intensity):
        self._left = Intensity(int(value)).value

    # endregion properties


class TblRelRight(TblRelLeft):
    """
    Relative Table size. Set table left margin using width as a percentage value.
    """

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.RIGHT)

    def _set_width_properties(self) -> None:
        page_width = self._prop_page_text_size.width
        self._set(self._props.right, 0)
        if self.prop_width.value == 100:
            self._set(self._props.width, page_width)
            self._set(self._props.left, 0)
            return
        width_factor = self.prop_width.value / 100
        width = round(page_width * width_factor)
        left = round(page_width - width)
        while left + width > page_width:
            # just in case rounding caused total to be more than page_width
            left = left - 1

        self._set(self._props.width, width)
        self._set(self._props.left, left)

    # endregion Overrides

    # region static methods
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TblRelRight:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRelRight:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRelRight:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelRight: ``TblRelRight`` Instance.
        """
        return TblRelRight._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblRelLeft._from_obj(cls, obj, **kwargs)
        return inst

    # endregion from_obj()
    # endregion static methods


class TblRelCenter(TblRelLeft):
    """
    Relative Table size. Set table left adn right margins using width as a percentage value.
    """

    # region Overrides
    def _post_init(self) -> None:
        self._set(self._props.hori_orient, HoriOrientation.CENTER)

    def _set_width_properties(self) -> None:
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
    def from_obj(cls, obj: object) -> TblRelCenter:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRelCenter:
        ...

    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TblRelCenter:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TblRelCenter: ``TblRelCenter`` Instance.
        """
        return TblRelCenter._from_obj(cls, obj, **kwargs)

    @staticmethod
    def _from_obj(cls: Type[_TTblAuto], obj: object, **kwargs) -> _TTblAuto:
        inst = TblRelLeft._from_obj(cls, obj, **kwargs)
        return inst

    # endregion from_obj()
    # endregion static methods


# endregion Table size in percentage units


class TableProperties(StyleMulti):
    """
    Frame Name options

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(self, name: str | None = None, size: TblPropObj | None = None) -> None:
        """
        Constructor

        Args:
            name (str, optional): Specifies frame name. Space are NOT allowed in names
            width (RelativeSize, AbsoluteSize optional): Specifies table Width.

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
        if not size is None:
            if not isinstance(size, TblAuto):
                raise ValueError(f'Class of "{type(size).__name__}" is not supported')
            size_obj = size.copy(_cattribs=self._get_tbl_cattribs())
            self._set_style("size", size_obj)

    # endregion Init

    # region Overrides

    def apply(self, obj: object, **kwargs) -> None:
        super().apply(obj, **kwargs)
        # for some reason setting Name property raises "UnknownPropertyException" when "setPropertyValue()" is used (Which Props.set() uses).
        # However, setting Name via setattr() works fine.
        # for this reason this class cancels setting of Name property and sets it via setattr() here.
        if self._has(self._props.name):
            if hasattr(obj, self._props.name):
                setattr(obj, self._props.name, self.prop_name)

    def on_property_setting(self, event_args: KeyValCancelArgs) -> None:
        if event_args.key == self._props.name:
            event_args.cancel = True
            event_args.handled = True
            # see bug specified in apply() method.
        super().on_property_setting(event_args)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = _get_default_tbl_services()
        return self._supported_services_values

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            return super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region internal methods
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
    def from_obj(cls: Type[_TTableProperties], obj: object) -> _TTableProperties:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTableProperties], obj: object, **kwargs) -> _TTableProperties:
        ...

    @classmethod
    def from_obj(cls: Type[_TTableProperties], obj: object, **kwargs) -> _TTableProperties:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Names: Instance that represents Frame Name options.
        """
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

        if rel:
            if hori == HoriOrientation.LEFT:
                size_obj = TblRelLeft.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.LEFT_AND_WIDTH:
                size_obj = TblRelFromLeft.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.RIGHT:
                size_obj = TblRelRight.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            else:
                size_obj = TblRelCenter.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
        else:
            if hori == HoriOrientation.FULL:
                size_obj = TblAuto.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.CENTER:
                size_obj = TblWidth.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.LEFT:
                size_obj = TblLeft.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.RIGHT:
                size_obj = TblRight.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            elif hori == HoriOrientation.LEFT_AND_WIDTH:
                size_obj = TblFromLeft.from_obj(obj, _cattribs=inst._get_tbl_cattribs())
            else:
                size_obj = TblManual.from_obj(obj, _cattribs=inst._get_tbl_cattribs())

        inst._set_style("size", size_obj)

        # prev, next not currently working
        return inst

    # endregion from_obj()

    # endregion static methods

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
    def prop_obj(self) -> TblPropObj | None:
        """Gets ``TblPropObj`` instance if exst."""
        try:
            return self._prop_inner_obj
        except AttributeError:
            val = self._get_style("size")
            if val is None:
                self._prop_inner_obj = None
            else:
                self._prop_inner_obj = val.style
        return self._prop_inner_obj

    @property
    def _props(self) -> TablePropertiesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _get_default_tbl_props()
        return self._props_internal_attributes

    # endregion properties
