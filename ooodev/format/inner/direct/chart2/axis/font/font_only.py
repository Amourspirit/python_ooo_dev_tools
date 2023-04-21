# region Import
from __future__ import annotations
from typing import Tuple
import uno
from ooodev.format.inner.direct.write.char.font.font_only import FontLang, FontOnly as CharFontOnly
from ooodev.units.unit_obj import UnitObj
from ooodev.units.unit_pt import UnitPT

# endregion Import

# for unnknown reason Chart DataPoint font size only responds when be set as an int
# even though the property is a float.
# If 10.5 is set for the font size it just gets ignored, this is true for all fractions.
# The fixe seems to be to force ints.


class FontOnly(CharFontOnly):
    """
    Character Font for a chart Datapoint.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        *,
        name: str | None = None,
        size: int | UnitObj | None = None,
        font_style: str | None = None,
        lang: FontLang | None = None,
    ) -> None:
        """
        Font options used in styles.

        Args:
            name (str, optional): This property specifies the name of the font style. It may contain more than one
                name separated by comma.
            size (int, UnitObj, optional): This value contains the size of the characters in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            font_style (str, optional): Font style name such as ``Bold``.
            lang (Lang, optional): Font Language
        """

        super().__init__(name=name, size=size, font_style=font_style, lang=lang)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Axis",)
        return self._supported_services_values

    @property
    def prop_size(self) -> UnitPT | None:
        """This value contains the size of the characters in point."""
        # for chart title this value is an int.
        return UnitPT(self._get(self._props.size))

    @prop_size.setter
    def prop_size(self, value: float | UnitObj | None) -> None:
        if value is None:
            self._remove(self._props.size)
            return
        try:
            self._set(self._props.size, round(value.get_value_pt()))
        except AttributeError:
            self._set(self._props.size, round(value))
