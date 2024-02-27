"""
Module for managing paragraph Text Flow.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, cast, Type, TypeVar, overload

from ooo.dyn.style.break_type import BreakType

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.direct.write.para.text_flow.breaks import Breaks
from ooodev.format.inner.direct.write.para.text_flow.hyphenation import Hyphenation
from ooodev.format.inner.direct.write.para.text_flow.flow_options import FlowOptions

_TTextFlow = TypeVar(name="_TTextFlow", bound="TextFlow")


class TextFlow(StyleMulti):
    """
    Paragraph Text Flow

    .. seealso::

        - :ref:`help_writer_format_direct_para_text_flow`

    .. versionadded:: 0.9.0
    """

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
            hy_end_chars (int, optional): characters at line end.
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

        See Also:

            - :ref:`help_writer_format_direct_para_text_flow`
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
            self._set_style("hyphenation", hy, *hy.get_attrs())  # type: ignore
        if brk.prop_has_attribs:
            self._set_style("breaks", brk, *brk.get_attrs())  # type: ignore
        if flo.prop_has_attribs:
            self._set_style("flow_options", flo, *flo.get_attrs())  # type: ignore

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TTextFlow], obj: Any) -> _TTextFlow: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TTextFlow], obj: Any, **kwargs) -> _TTextFlow: ...

    @classmethod
    def from_obj(cls: Type[_TTextFlow], obj: Any, **kwargs) -> _TTextFlow:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            TextFlow: ``TextFlow`` instance that represents ``obj`` Indents and spacing.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        hy = Hyphenation.from_obj(obj)
        if hy.prop_has_attribs:
            inst._set_style("hyphenation", hy, *hy.get_attrs())  # type: ignore
        brk = Breaks.from_obj(obj)
        if brk.prop_has_attribs:
            inst._set_style("breaks", brk, *brk.get_attrs())  # type: ignore
        flo = FlowOptions.from_obj(obj)
        if flo.prop_has_attribs:
            inst._set_style("flow_options", flo, *flo.get_attrs())  # type: ignore

        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

    @property
    def prop_inner_hyphenation(self) -> Hyphenation | None:
        """Gets Hyphenation instance"""
        try:
            return self._direct_inner_hy
        except AttributeError:
            self._direct_inner_hy = cast(Hyphenation, self._get_style_inst("hyphenation"))
        return self._direct_inner_hy

    @property
    def prop_inner_breaks(self) -> Breaks | None:
        """Gets Breaks instance"""
        try:
            return self._direct_inner_breaks
        except AttributeError:
            self._direct_inner_breaks = cast(Breaks, self._get_style_inst("breaks"))
        return self._direct_inner_breaks

    @property
    def prop_inner_flow_options(self) -> FlowOptions | None:
        """Gets Flow Options instance"""
        try:
            return self._direct_inner_fo
        except AttributeError:
            self._direct_inner_fo = cast(FlowOptions, self._get_style_inst("flow_options"))
        return self._direct_inner_fo

    @property
    def default(self: _TTextFlow) -> _TTextFlow:
        """Gets ``TextFlow`` default."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            if self.prop_inner_hyphenation is None:
                hy = Hyphenation().default
            else:
                hy = self.prop_inner_hyphenation.default
            if self.prop_inner_breaks is None:
                brk = Breaks().default
            else:
                brk = self.prop_inner_breaks.default
            if self.prop_inner_flow_options is None:
                flo = FlowOptions().default
            else:
                flo = self.prop_inner_flow_options.default
            tf = self.__class__()
            tf._set_style("hyphenation", hy, *hy.get_attrs())  # type: ignore
            tf._set_style("breaks", brk, *brk.get_attrs())  # type: ignore
            tf._set_style("flow_options", flo, *flo.get_attrs())  # type: ignore
            tf._is_default_inst = True
            self._default_inst = tf
        return self._default_inst

    # endregion properties
