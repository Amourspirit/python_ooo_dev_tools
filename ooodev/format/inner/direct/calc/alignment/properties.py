# region Import
from __future__ import annotations
from typing import cast, overload
from typing import Any, Tuple, Type, TypeVar
from enum import Enum
from ooo.dyn.text.writing_mode2 import WritingMode2

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.cell_text_properties_props import CellTextPropertiesProps

# endregion Import

_TProperties = TypeVar(name="_TProperties", bound="Properties")


class TextDirectionKind(Enum):
    """
    Describes different writing directions
    """

    LR_TB = WritingMode2.LR_TB
    """
    Left-to-right (LTR)
    
    Text within lines is written left-to-right.
    Lines and blocks are placed top-to-bottom.
    Typically, this is the writing mode for normal ``alphabetic`` text.
    """
    RL_TB = WritingMode2.RL_TB
    """
    Right-to-left (RTL).
    
    text within a line are written right-to-left.
    Lines and blocks are placed top-to-bottom.
    Typically, this writing mode is used in Arabic and Hebrew text.
    """

    PAGE = WritingMode2.PAGE  # use superordinate object settings
    """
    Use super-ordinate object settings
    
    Obtain writing mode from the current page.
    May not be used in page styles.
    """

    def __int__(self) -> int:
        return self.value


class Properties(StyleBase):
    """
    Text Properties

    .. seealso::

        - :ref:`help_calc_format_direct_cell_alignment`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        wrap_auto: bool | None = None,
        hyphen_active: bool | None = None,
        shrink_to_fit: bool | None = None,
        direction: TextDirectionKind | None = None,
    ) -> None:
        """
        Constructor

        Args:
            wrap_auto (bool, optional): Specifies wrap text automatically.
            hyphen_active (bool, optional): Specifies hyphenation active.
            shrink_to_fit (bool, optional): Specifies if text will shrink to cell.
            direction (TextDirectionKind, optional): Specifies Text Direction.

        Returns:
            None:

        Note:
            When ``wrap_auto`` is ``True`` ``shrink_to_fit`` is not used.

        See Also:
            - :ref:`help_calc_format_direct_cell_alignment`
        """

        super().__init__()
        if wrap_auto is not None:
            self.prop_wrap_auto = wrap_auto
        if hyphen_active is not None:
            self.prop_hyphen_active = hyphen_active
        if shrink_to_fit is not None:
            self.prop_shrink_to_fit = shrink_to_fit
        if direction is not None:
            self.prop_direction = direction

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CellStyle", "com.sun.star.table.CellProperties")
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
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
    def from_obj(cls: Type[_TProperties], obj: Any) -> _TProperties: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TProperties], obj: Any, **kwargs) -> _TProperties: ...

    @classmethod
    def from_obj(cls: Type[_TProperties], obj: Any, **kwargs) -> _TProperties:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Properties: Instance that represents text property options.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')
        for prop in inst._props:
            if prop:
                val = mProps.Props.get(obj, prop, None)
                if val is not None:
                    inst._set(prop, val)
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
            self._format_kind_prop = FormatKind.CELL
        return self._format_kind_prop

    @property
    def prop_wrap_auto(self) -> bool | None:
        """
        Gets/Sets If text is wrapped automatically.
        """
        return self._get(self._props.wrapped)

    @prop_wrap_auto.setter
    def prop_wrap_auto(self, value: bool | None):
        if value is None:
            self._remove(self._props.wrapped)
            return
        self._set(self._props.wrapped, value)

    @property
    def prop_hyphen_active(self) -> bool | None:
        """
        Gets/Sets If text is hyphenation is active.
        """
        return self._get(self._props.hyphen)

    @prop_hyphen_active.setter
    def prop_hyphen_active(self, value: bool | None):
        if value is None:
            self._remove(self._props.hyphen)
            return
        self._set(self._props.hyphen, value)

    @property
    def prop_shrink_to_fit(self) -> bool | None:
        """
        Gets/Sets If text shrinks to cell size.
        """
        return self._get(self._props.shrink)

    @prop_shrink_to_fit.setter
    def prop_shrink_to_fit(self, value: bool | None):
        if value is None:
            self._remove(self._props.shrink)
            return
        self._set(self._props.shrink, value)

    @property
    def prop_direction(self) -> TextDirectionKind | None:
        """
        Gets/Sets Text Direction Kind.
        """
        pv = cast(int, self._get(self._props.mode))
        return None if pv is None else TextDirectionKind(pv)

    @prop_direction.setter
    def prop_direction(self, value: TextDirectionKind | None):
        if value is None:
            self._remove(self._props.mode)
            return
        self._set(self._props.mode, int(value))

    @property
    def _props(self) -> CellTextPropertiesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = CellTextPropertiesProps(
                mode="WritingMode", wrapped="IsTextWrapped", hyphen="ParaIsHyphenation", shrink="ShrinkToFit"
            )
        return self._props_internal_attributes

    # endregion Properties
