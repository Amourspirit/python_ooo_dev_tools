"""
Module for managing paragraph alignment.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar
from enum import Enum

from ooo.dyn.style.paragraph_adjust import ParagraphAdjust
from ooo.dyn.text.paragraph_vert_align import ParagraphVertAlignEnum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.direct.write.para.align.writing_mode import WritingMode
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps

_TAlignment = TypeVar(name="_TAlignment", bound="Alignment")


class LastLineKind(Enum):
    """Last Line Alignment kind"""

    START = 0
    """Align Start"""
    JUSTIFY = 2
    """Align justified"""
    CENTER = 3
    """Align Center"""


class Alignment(StyleMulti):
    """
    Paragraph Alignment

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_alignment`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        align: ParagraphAdjust | None = None,
        align_vert: ParagraphVertAlignEnum | None = None,
        txt_direction: WritingMode | None = None,
        align_last: LastLineKind | None = None,
        expand_single_word: bool | None = None,
        snap_to_grid: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            align (ParagraphAdjust, optional): Determines horizontal alignment of a paragraph.
            align_vert (ParagraphVertAlignEnum, optional): Determines vertical alignment of a paragraph.
            text_direction (WritingMode, optional): Determines the text direction.
            align_last (LastLineKind, optional): Determines the adjustment of the last line.
            expand_single_word (bool, optional): Determines if single words are stretched.
                It is only valid if ``align`` and ``align_last`` are also valid.
            snap_to_grid (bool, optional): Determines snap to text grid (if active).

        Returns:
            None:

        Hint:
            - ``ParagraphAdjust`` can be imported from ``ooo.dyn.style.paragraph_adjust``
            - ``ParagraphVertAlignEnum`` can be imported from ``ooo.dyn.text.paragraph_vert_align``
            - ``LastLineKind`` can be imported from ``ooodev.format.writer.direct.para.alignment``
            - ``WritingMode`` can be imported from ``ooodev.format.inner.direct.write.para.align``
            - ``WritingMode2Enum`` can be imported from ``oo.dyn.text.writing_mode2``.

        See Also:

            - :ref:`help_writer_format_direct_para_alignment`
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if align is not None:
            # ParagraphAdjust.STRETCH seems to be the same as LEFT
            init_vals["ParaAdjust"] = align

        if align_vert is not None:
            init_vals["ParaVertAlignment"] = align_vert.value

        if align_last is not None:
            init_vals["ParaLastLineAdjust"] = align_last.value

        # SnapToGrid: could not find what service this property is part of, may not be any.
        if snap_to_grid is not None:
            init_vals["SnapToGrid"] = snap_to_grid

        if expand_single_word is not None:
            init_vals["ParaExpandSingleWord"] = expand_single_word

        super().__init__(**init_vals)
        if txt_direction is not None:
            self._set_style("txt_direction", txt_direction, *txt_direction.get_attrs())

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

        set_prop("ParaVertAlignment", inst)
        set_prop("ParaLastLineAdjust", inst)
        set_prop("ParaExpandSingleWord", inst)

        # LibreOffice Writer converts ParagraphAdjust to an int value
        adjust = cast(int, mProps.Props.get(obj, "ParaAdjust"))
        inst._set("ParaAdjust", cls.convert_int_to_paragraph_adjust(adjust))
        try:
            # SnapToGrid is not part of any known service
            snap = mProps.Props.get(obj, "SnapToGrid")
            inst._set("SnapToGrid", snap)
        except mEx.PropertyNotFoundError as e:
            mLo.Lo.print("Alignment.from_obj(), SnapToGrid property not found")
            mLo.Lo.print(f"  {e}")

        try:
            txt_dir = WritingMode.from_obj(obj)
            inst._set_style("txt_direction", txt_dir, *txt_dir.get_attrs())
        except Exception:
            mLo.Lo.print("Alignment.from_obj(): unable to set txt_direction style")
        inst.set_update_obj(obj)
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

    def fmt_align_vert(self: _TAlignment, value: ParagraphVertAlignEnum | None) -> _TAlignment:
        """
        Gets copy of instance with vertical alignment set or removed

        Args:
            value (ParagraphVertAlignEnum | None): Alignment value

        Returns:
            Alignment: Alignment instance
        """
        cp = self.copy()
        cp.prop_align_vert = value
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

    def fmt_snap_to_grid(self: _TAlignment, value: bool | None) -> _TAlignment:
        """
        Gets copy of instance with snap to grid set or removed

        Args:
            value (LastLineKind | None): Snap to grid value

        Returns:
            Alignment: Alignment instance
        """
        cp = self.copy()
        cp.prop_snap_to_grid = value
        return cp

    def fmt_txt_direction(self: _TAlignment, value: WritingMode | None) -> _TAlignment:
        """
        Gets copy of instance with vertical alignment set or removed

        Args:
            value (ParagraphVertAlignEnum | None): Alignment value

        Returns:
            Alignment: Alignment instance
        """
        cp = self.copy()
        if value is None:
            self._remove_style("txt_direction")
        else:
            self._set_style("txt_direction", value)
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def snap_to_grid(self: _TAlignment) -> _TAlignment:
        """Gets copy of instance with snap to grid set"""
        al = self.copy()
        al.prop_snap_to_grid = True
        return al

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
            self._format_kind_prop = FormatKind.PARA
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
    def prop_align_vert(self) -> ParagraphVertAlignEnum | None:
        """Gets/Sets vertical alignment of a paragraph."""
        pv = cast(int, self._get("ParaVertAlignment"))
        return None if pv is None else ParagraphVertAlignEnum(pv)

    @prop_align_vert.setter
    def prop_align_vert(self, value: ParagraphVertAlignEnum | None):
        if value is None:
            self._remove("ParaVertAlignment")
            return
        self._set("ParaVertAlignment", value)

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
    def prop_snap_to_grid(self) -> bool | None:
        """Gets/Sets snap to text grid (if active)."""
        # SnapToGrid is not part of any know service
        return self._get("SnapToGrid")

    @prop_snap_to_grid.setter
    def prop_snap_to_grid(self, value: bool | None):
        if value is None:
            self._remove("SnapToGrid")
            return
        self._set("SnapToGrid", value)

    @property
    def prop_inner_mode(self) -> WritingMode | None:
        """Gets Writing Mode (``txt_direction``) instance if exist."""
        try:
            return self._direct_inner_mode
        except AttributeError:
            self._direct_inner_mode = cast(WritingMode, self._get_style_inst("txt_direction"))
        return self._direct_inner_mode

    @property
    def default(self: _TAlignment) -> _TAlignment:
        """Gets Alignment default."""
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            if self.prop_inner_mode is None:
                mode = WritingMode(_cattribs=self._get_internal_cattribs()).default  # type: ignore
            else:
                mode = self.prop_inner_mode.default
            self._default_inst = self.__class__(
                align=ParagraphAdjust.LEFT,
                align_vert=ParagraphVertAlignEnum.AUTOMATIC,
                txt_direction=mode,
                align_last=LastLineKind.START,
                expand_single_word=False,
                snap_to_grid=True,
                _cattribs=self._get_internal_cattribs(),  # type: ignore
            )
            # pylint: disable=protected-access
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
