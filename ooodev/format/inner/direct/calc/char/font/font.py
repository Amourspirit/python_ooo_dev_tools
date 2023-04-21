from __future__ import annotations
import uno
from ooodev.format.inner.direct.write.char.font.font import Font as CharFont


class Font(CharFont):
    """
    Calc Character Font

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties such as ``bold``, ``italic``, ``underline`` can be chained together.

    Example:

        .. code-block:: python

            # chaining fonts together to add new properties
            ft = Font().bold.italic.underline

            ft_color = Font().style_color(CommonColor.GREEN).style_bg_color(CommonColor.LIGHT_BLUE)

    .. versionadded:: 0.9.4
    """

    pass
