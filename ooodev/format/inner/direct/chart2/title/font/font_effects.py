# region Imports
from __future__ import annotations
from typing import Tuple
import uno
from ooodev.format.inner.direct.write.char.font.font_effects import FontEffects as CharFontEffects


# endregion Imports


class FontEffects(CharFontEffects):
    """
    Character Font Effects for a chart title.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties can be chained together.

    .. versionadded:: 0.9.4
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.beans.PropertySet",
                "com.sun.star.style.CharacterProperties",
            )
        return self._supported_services_values
