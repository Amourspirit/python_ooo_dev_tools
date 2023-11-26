# region Import
from __future__ import annotations
import uno
from ooodev.format.inner.direct.write.char.font.font_only import FontOnly as CharFontOnly, FontLang
from ooodev.units import UnitT

# endregion Import


class FontOnly(CharFontOnly):
    """
    Character Font for a chart legend.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_legend_font_only`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        *,
        name: str | None = None,
        size: float | UnitT | None = None,
        font_style: str | None = None,
        lang: FontLang | None = None,
    ) -> None:
        """
        Font options used in styles.

        Args:
            name (str, optional): This property specifies the name of the font style. It may contain more than one
                name separated by comma.
            size (float, UnitT, optional): This value contains the size of the characters in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            font_style (str, optional): Font style name such as ``Bold``.
            lang (Lang, optional): Font Language

        Returns:
            None:

        See Also:

            - :ref:`help_chart2_format_direct_legend_font_only`
        """
        super().__init__(name=name, size=size, font_style=font_style, lang=lang)
