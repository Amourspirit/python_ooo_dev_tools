"""
Base Class for Para Style.

.. versionadded:: 0.9.0
"""
from __future__ import annotations

import uno
from ...kind.format_kind import FormatKind
from ...style_base import StyleModifyMulti
from ...writer.style.para.kind import StyleParaKind as StyleParaKind


class ParaStyleBaseMulti(StyleModifyMulti):
    """
    Para Style Base Multi

    .. versionadded:: 0.9.0
    """

    def _get_style_family_name(self) -> str:
        return "ParagraphStyles"

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA | FormatKind.STYLE
