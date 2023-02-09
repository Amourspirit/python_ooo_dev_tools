"""
Module for Shadow format (``ShadowFormat``) struct.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import Dict, Tuple, cast, overload

from ....events.event_singleton import _Events
from ....meta.static_prop import static_prop
from ....utils import props as mProps
from ....utils.color import Color
from ....utils.color import CommonColor
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent
from ....utils.unit_convert import UnitConvert, Length
from ....utils.type_var import T

import uno
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


# endregion imports
class ShadowStruct(StyleBase):
    """
    Shadow struct

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = CommonColor.GRAY,
        transparent: bool = False,
        width: float = 1.76,
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow. Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (Color, optional):contains the color value of the shadow. Defaults to ``CommonColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, optional): contains the size of the shadow (in mm units). Defaults to ``1.76``.

        Raises:
            ValueError: If ``color`` or ``width`` are less than zero.
        """
        if color < 0:
            raise ValueError("color must be a positive number")
        if width < 0:
            raise ValueError("Width must be a postivie number")

        self._location = location
        self._color = color
        self._transparent = transparent
        self._width: int = round(UnitConvert.convert(num=width, frm=Length.MM, to=Length.MM100))

        super().__init__()

    # endregion init

    # region methods
    def __eq__(self, other: object) -> bool:
        s2: ShadowFormat = None
        if isinstance(other, ShadowStruct):
            s2 = other.get_shadow_format()
        elif getattr(other, "typeName", None) == "com.sun.star.table.ShadowFormat":
            s2 = other
        if s2:
            s1 = self.get_shadow_format()
            return (
                s1.Color == s2.Color
                and s1.IsTransparent == s2.IsTransparent
                and s1.Location == s2.Location
                and s1.ShadowWidth == s2.ShadowWidth
            )
        return False

    def get_shadow_format(self) -> ShadowFormat:
        """
        Gets Shadow format for instance.

        Returns:
            ShadowFormat: Shadow Format
        """
        return ShadowFormat(
            Location=self._location,
            ShadowWidth=self._width,
            IsTransparent=self._transparent,
            Color=self._color,
        )

    def _supported_services(self) -> Tuple[str, ...]:
        return ()

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _get_property_name(self) -> str:
        return "ShadowFormat"

    def get_attrs(self) -> Tuple[str, ...]:
        """
        Gets the attributes that are slated for change in the current instance

        Returns:
            Tuple(str, ...): Tuple of attribures
        """
        return (self._get_property_name(),)

    def copy(self: T) -> T:
        nu = super(ShadowStruct, self.__class__).__new__(self.__class__)
        nu.__init__(
            location=self.prop_width, color=self.prop_color, transparent=self.prop_transparent, width=self.prop_width
        )
        if self._dv:
            nu._update(self._dv)
        return nu

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    @overload
    def apply(self, obj: object, keys: Dict[str, str]) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object that contains a ``ShadowFormat`` property.
            keys: (Dict[str, str], optional): key map for properties.
                Can be ``prop`` which defaults to ``ShadowFormat``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_obj("apply()")
            return

        keys = {"prop": self._get_property_name()}
        if "keys" in kwargs:
            keys.update(kwargs["keys"])

        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return

        shadow = self.get_shadow_format()
        mProps.Props.set(obj, **{keys["prop"]: shadow})
        eargs = EventArgs.from_args(cargs)
        self.on_applied(eargs)
        _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)

    # endregion apply()

    @classmethod
    def from_obj(cls, obj: object) -> ShadowStruct:
        """
        Gets instance from object

        Args:
            obj (object): UNO object

        Returns:
            Shadow: Instance from object
        """

        # this nu is only used to get Property Name
        nu = super(ShadowStruct, cls).__new__(cls)
        nu.__init__()

        shadow = cast(ShadowFormat, mProps.Props.get(obj, nu._get_property_name()))
        if shadow is None:
            return cls.empty.copy()
        width = UnitConvert.convert(num=shadow.ShadowWidth, frm=Length.MM100, to=Length.MM)

        nu = super(ShadowStruct, cls).__new__(cls)
        nu.__init__(location=shadow.Location, color=shadow.Color, transparent=shadow.IsTransparent, width=width)
        return nu

    @classmethod
    def from_shadow(cls, shadow: ShadowFormat) -> ShadowStruct:
        """
        Gets an instance

        Args:
            shadow (ShadowFormat): Shadow Format

        Returns:
            Shadow: Instance representing ``shadow``.
        """
        width = UnitConvert.convert(num=shadow.ShadowWidth, frm=Length.MM100, to=Length.MM)
        nu = super(ShadowStruct, cls).__new__(cls)
        nu.__init__(location=shadow.Location, color=shadow.Color, transparent=shadow.IsTransparent, width=width)
        return nu

    # endregion methods

    # region style methods
    def fmt_location(self, value: ShadowLocation) -> ShadowStruct:
        """
        Gets a copy of instance with location set

        Args:
            value (ShadowLocation): Shadow location value

        Returns:
            Shadow: Shadow with location set
        """
        cp = self.copy()
        cp.prop_location = value
        return cp

    def fmt_color(self, value: Color) -> ShadowStruct:
        """
        Gets a copy of instance with color set

        Args:
            value (Color): color value

        Returns:
            Shadow: Shadow with color set
        """
        cp = self.copy()
        cp.prop_color = value
        return cp

    def fmt_transparent(self, value: bool) -> ShadowStruct:
        """
        Gets a copy of instance with transparency set

        Args:
            value (bool): transparency value

        Returns:
            Shadow: Shadow with transparency set
        """
        cp = self.copy()
        cp.prop_transparent = value
        return cp

    def fmt_width(self, value: float) -> ShadowStruct:
        """
        Gets a copy of instance with width set

        Args:
            value (float): width value

        Returns:
            Shadow: Shadow with width set
        """
        cp = self.copy()
        cp.prop_width = value
        return cp

    # endregion style methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.STRUCT

    @property
    def prop_location(self) -> ShadowLocation:
        """Gets the location of the shadow."""
        return self._location

    @prop_location.setter
    def prop_location(self, value: ShadowLocation) -> None:
        self._location = value

    @property
    def prop_color(self) -> Color:
        """Gets the color value of the shadow."""
        return self._color

    @prop_color.setter
    def prop_color(self, value: Color) -> None:
        self._color = value

    @property
    def prop_transparent(self) -> bool:
        """Gets transparent value"""
        return self._transparent

    @prop_transparent.setter
    def prop_transparent(self, value: bool) -> None:
        self._transparent = value

    @property
    def prop_width(self) -> float:
        """Gets the size of the shadow (in mm units)"""
        return UnitConvert.convert(num=self._width, frm=Length.MM100, to=Length.MM)

    @prop_width.setter
    def prop_width(self, value: float) -> None:
        self._width = round(UnitConvert.convert(num=value, frm=Length.MM, to=Length.MM100))

    @static_prop
    def empty() -> ShadowStruct:  # type: ignore[misc]
        """Gets empty Shadow. Static Property. when style is applied it remove any shadow."""
        try:
            return ShadowStruct._EMPTY_INST
        except AttributeError:
            ShadowStruct._EMPTY_INST = ShadowStruct(
                location=ShadowLocation.NONE, transparent=False, color=8421504, width=1.76
            )
            ShadowStruct._EMPTY_INST._is_default_inst = True
        return ShadowStruct._EMPTY_INST

    # endregion Properties
