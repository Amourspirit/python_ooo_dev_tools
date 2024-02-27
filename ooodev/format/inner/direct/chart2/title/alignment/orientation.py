from __future__ import annotations
from typing import cast, TypeVar, Tuple
import uno

from ooodev.format.inner.common.props.title_alignment_orientation_props import TitleAlignmentOrientationProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.angle import Angle

_TOrientation = TypeVar("_TOrientation", bound="Orientation")


class Orientation(StyleBase):
    """
    Title Text orientation.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_title_alignment`

    .. versionadded:: 0.9.4
    """

    def __init__(self, angle: int | Angle | None = None, vertical: bool | None = None) -> None:
        """
        Constructor

        Args:
            angle (int, Angle, optional): Rotation in degrees of the text.
            vertical (bool, optional): Specifies if the text is vertically stacked.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_title_alignment`
        """
        super().__init__()
        self.prop_angle = angle
        self.prop_vertical = vertical

    # region style methods
    def fmt_angle(self: _TOrientation, value: int | Angle | None) -> _TOrientation:
        """
        Gets new instance with the rotation set or removed.

        Args:
            value (int | Angle | None): The rotation in degrees, ``None`` to remove.

        Returns:
            _TOrientation: The new instance.
        """
        cp = self.copy()
        cp.prop_angle = value
        return cp

    def fmt_vertical(self: _TOrientation, value: bool | None) -> _TOrientation:
        """
        Gets new instance with the vertical orientation set or removed.

        Args:
            value (bool | None): ``True`` or ``False`` to set vertical orientation, ``None`` to remove.

        Returns:
            _TOrientation: The new instance.
        """
        cp = self.copy()
        cp.prop_vertical = value
        return cp

    # endregion style methods

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Title",)
        return self._supported_services_values

    # endregion overrides

    # region Style Properties
    @property
    def vertical(self: _TOrientation) -> _TOrientation:
        """Gets new instance with the vertical orientation set."""
        cp = self.copy()
        cp.prop_vertical = True
        return cp

    # endregion Style Properties

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def prop_angle(self) -> Angle | None:
        """Gets/Sets the rotation of the text."""
        pv = cast(int, self._get(self._props.angle))
        return None if pv is None else Angle(pv)

    @prop_angle.setter
    def prop_angle(self, value: int | Angle | None) -> None:
        if value is None:
            self._remove(self._props.angle)
            return
        val = Angle(int(value))
        self._set(self._props.angle, val.value)

    @property
    def prop_vertical(self) -> bool | None:
        """Gets/Sets if the text is vertically stacked."""
        pv = cast(bool, self._get(self._props.vertical))
        return None if pv is None else pv

    @prop_vertical.setter
    def prop_vertical(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.vertical)
            return
        self._set(self._props.vertical, value)

    @property
    def _props(self) -> TitleAlignmentOrientationProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = TitleAlignmentOrientationProps(
                angle="TextRotation", vertical="StackCharacters"
            )
        return self._props_internal_attributes

    # endregion Properties
