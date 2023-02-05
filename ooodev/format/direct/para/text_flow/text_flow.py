"""
Modele for managing paragraph Text Flow.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple

from .....meta.static_prop import static_prop
from ....style_base import StyleMulti
from ....kind.format_kind import FormatKind
from .breaks import Breaks
from .hyphenation import Hyphenation
from .flow_options import FlowOptions

from ooo.dyn.style.break_type import BreakType as BreakType


class TextFlow(StyleMulti):
    """
    Paragraph Text Flow

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        *,
        hy_auto: bool | None = None,
        hy_no_caps: bool | None = None,
        hy_start_chars: int | None = None,
        hy_end_chars: int | None = None,
        hy_max: int | None = None,
        bk_type: BreakType | None = None,
        bk_style: str | None = None,
        bk_num: int | None = None,
        op_orphans: int | None = None,
        op_widows: int | None = None,
        op_keep: bool | None = None,
        op_no_split: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            hy_auto (bool, optional): Hyphenate automatically.
            hy_no_caps (bool, optional): Don't hyphenate word in caps.
            hy_start_chars (int, optional): Characters at line begin.
            hy_end_chars (int, optional): charactors at line end.
            hy_max (int, optional): Maximum consecutive hyphenated lines.
            bk_type (Any, optional): Break type.
            bk_style (str, optional): Style to apply to break.
            bk_num (int, optional): Page number to apply to break.
            op_orphans (int, optional): Number of Orphan Control Lines.
            op_widows (int, optional): Number Widow Control Lines.
            op_keep (bool, optional): Keep with next paragraph.
            op_no_split (bool, optional): Do not split paragraph.
        Returns:
            None:

        Note:
            Arguments starting with ``hy_`` are for hyphenation

            Arguments starting with ``bk_`` are for Breaks

            Arguments starting with ``op_`` are for Flow options
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        hy = Hyphenation(
            auto=hy_auto, no_caps=hy_no_caps, start_chars=hy_start_chars, end_chars=hy_end_chars, max=hy_max
        )

        brk = Breaks(type=bk_type, style=bk_style, num=bk_num)

        flo = FlowOptions(orphans=op_orphans, widows=op_widows, keep=op_keep, no_split=op_no_split)

        super().__init__(**init_vals)
        if hy.prop_has_attribs:
            self._set_style("hyphenation", hy, *hy.get_attrs())
        if brk.prop_has_attribs:
            self._set_style("breaks", brk, *brk.get_attrs())
        if flo.prop_has_attribs:
            self._set_style("flow_options", flo, *flo.get_attrs())

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    @staticmethod
    def from_obj(obj: object) -> TextFlow:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            TextFlow: ``TextFlow`` instance that represents ``obj`` Indents and spacing.
        """
        ids = TextFlow()
        hy = Hyphenation.from_obj(obj)
        if hy.prop_has_attribs:
            ids._set_style("hyphenation", hy, *hy.get_attrs())
        brk = Breaks.from_obj(obj)
        if brk.prop_has_attribs:
            ids._set_style("breaks", brk, *brk.get_attrs())
        flo = FlowOptions.from_obj(obj)
        if flo.prop_has_attribs:
            ids._set_style("flow_options", flo, *flo.get_attrs())
        return ids

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @static_prop
    def default() -> TextFlow:  # type: ignore[misc]
        """Gets ``TextFlow`` default. Static Property."""
        if TextFlow._DEFAULT is None:
            hy = Hyphenation.default
            brk = Breaks.default
            flo = FlowOptions.default
            tf = TextFlow()
            tf._set_style("hyphenation", hy, *hy.get_attrs())
            tf._set_style("breaks", brk, *brk.get_attrs())
            tf._set_style("flow_options", flo, *flo.get_attrs())
            TextFlow._DEFAULT = tf
        return TextFlow._DEFAULT

    # endregion properties
