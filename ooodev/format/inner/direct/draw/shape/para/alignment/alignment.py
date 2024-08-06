"""
Module for managing shape paragraph alignment.

.. versionadded:: 0.17.8
"""

from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar
import uno
from ooo.dyn.style.paragraph_adjust import ParagraphAdjust as ParagraphAdjust
from ooo.dyn.text.paragraph_vert_align import ParagraphVertAlignEnum as ParagraphVertAlignEnum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.direct.write.para.align.alignment import LastLineKind

_TAlignment = TypeVar("_TAlignment", bound="Alignment")


class Alignment(StyleBase):
    """
    Shape Paragraph Alignment

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.17.8
    """

    # region init

    def __init__(
        self,
        *,
        align: ParagraphAdjust | None = None,
        align_last: LastLineKind | None = None,
        expand_single_word: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            align (ParagraphAdjust, optional): Determines horizontal alignment of a paragraph.
            align_last (LastLineKind, optional): Determines the adjustment of the last line.
            expand_single_word (bool, optional): Determines if single words are stretched.
                It is only valid if ``align`` and ``align_last`` are also valid.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if align is not None:
            # ParagraphAdjust.STRETCH seems to be the same as LEFT
            init_vals["ParaAdjust"] = align

        if align_last is not None:
            init_vals["ParaLastLineAdjust"] = align_last.value

        if expand_single_word is not None:
            init_vals["ParaExpandSingleWord"] = expand_single_word

        super().__init__(**init_vals)

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

    def _on_modifying(self, source: Any, event_args: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event_args)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies alignment to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    # region static methods

    @staticmethod
    def convert_int_to_paragraph_adjust(num: int) -> ParagraphAdjust:
        """
        Converts integer value to ``ParagraphAdjust``.

        When Writer saves ``ParagraphAdjust`` value into ``ParaAdjust`` it converts it to a integer value.

        Args:
            num (int): Number to convert

        Returns:
            ParagraphAdjust: Number as ``ParagraphAdjust``
        """
        if num == 0:
            return ParagraphAdjust.LEFT
        if num == 1:
            return ParagraphAdjust.RIGHT
        if num == 2:
            return ParagraphAdjust.BLOCK
        return ParagraphAdjust.CENTER if num == 3 else ParagraphAdjust.STRETCH

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TAlignment], obj: Any) -> _TAlignment: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAlignment], obj: Any, **kwargs) -> _TAlignment: ...

    @classmethod
    def from_obj(cls: Type[_TAlignment], obj: Any, **kwargs) -> _TAlignment:
        """
        Gets Padding instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Alignment: Alignment that represents ``obj`` alignment.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, align: Alignment):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                align._set(key, val)

        set_prop("ParaLastLineAdjust", inst)
        set_prop("ParaExpandSingleWord", inst)

        # LibreOffice Writer converts ParagraphAdjust to an int value
        adjust = cast(int, mProps.Props.get(obj, "ParaAdjust"))
        inst._set("ParaAdjust", cls.convert_int_to_paragraph_adjust(adjust))
        return inst

    # endregion from_obj()

    # endregion static methods

    # endregion methods

    # region style methods
    def fmt_align(self: _TAlignment, value: ParagraphAdjust | None) -> _TAlignment:
        """
        Gets copy of instance with horizontal alignment set or removed

        Args:
            value (ParagraphAdjust | None): Alignment value

        Returns:
            Alignment: Alignment instance
        """
        cp = self.copy()
        cp.prop_align = value
        return cp

    def fmt_align_last(self: _TAlignment, value: LastLineKind | None) -> _TAlignment:
        """
        Gets copy of instance with align last set or removed

        Args:
            value (LastLineKind | None): Align last value

        Returns:
            Alignment: Alignment instance
        """
        cp = self.copy()
        cp.prop_align_last = value
        return cp

    def fmt_expand_single_word(self: _TAlignment, value: bool | None) -> _TAlignment:
        """
        Gets copy of instance with expand single word set or removed

        Args:
            value (LastLineKind | None): Expand single word value

        Returns:
            Alignment: Alignment instance
        """
        cp = self.copy()
        cp.prop_expand_single_word = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def expand_single_word(self: _TAlignment) -> _TAlignment:
        """Gets copy of instance with expand single word set"""
        al = self.copy()
        al.prop_expand_single_word = True
        return al

    @property
    def justified(self: _TAlignment) -> _TAlignment:
        """Gets copy of instance with align set to block"""
        al = self.copy()
        al.prop_align = ParagraphAdjust.BLOCK
        return al

    @property
    def align_center(self: _TAlignment) -> _TAlignment:
        """Gets copy of instance with align set to center"""
        al = self.copy()
        al.prop_align = ParagraphAdjust.CENTER
        return al

    @property
    def align_left(self: _TAlignment) -> _TAlignment:
        """Gets copy of instance with align set to left"""
        al = self.copy()
        al.prop_align = ParagraphAdjust.LEFT
        return al

    @property
    def align_right(self: _TAlignment) -> _TAlignment:
        """Gets copy of instance with align set to left"""
        al = self.copy()
        al.prop_align = ParagraphAdjust.RIGHT
        return al

    # endregion Style Properties

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA | FormatKind.SHAPE
        return self._format_kind_prop

    @property
    def prop_align(self) -> ParagraphAdjust | None:
        """Gets/Sets horizontal alignment of a paragraph."""
        return self._get("ParaAdjust")

    @prop_align.setter
    def prop_align(self, value: ParagraphAdjust | None):
        if value is None:
            self._remove("ParaAdjust")
            return
        self._set("ParaAdjust", value)

    @property
    def prop_align_last(self) -> LastLineKind | None:
        """Gets/Sets the adjustment of the last line."""
        pv = cast(int, self._get("ParaLastLineAdjust"))
        return None if pv is None else LastLineKind(pv)

    @prop_align_last.setter
    def prop_align_last(self, value: LastLineKind | None):
        if value is None:
            self._remove("ParaLastLineAdjust")
            return
        self._set("ParaLastLineAdjust", value)

    @property
    def prop_expand_single_word(self) -> bool | None:
        """
        Gets/Sets Determines if single words are stretched.

        It is only valid if ``prop_align`` and ``prop_align_last`` are also valid.
        """
        return self._get("ParaExpandSingleWord")

    @prop_expand_single_word.setter
    def prop_expand_single_word(self, value: bool | None):
        if value is None:
            self._remove("ParaExpandSingleWord")
            return
        self._set("ParaExpandSingleWord", value)

    @property
    def default(self: _TAlignment) -> _TAlignment:
        """Gets Alignment default."""
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(
                align=ParagraphAdjust.LEFT,
                align_last=LastLineKind.START,
                expand_single_word=False,
                _cattribs=self._get_internal_cattribs(),  # type: ignore
            )
            # pylint: disable=protected-access
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
