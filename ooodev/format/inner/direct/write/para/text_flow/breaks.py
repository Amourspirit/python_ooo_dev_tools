"""
Module for managing paragraph breaks.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import Any, Dict, Tuple, overload, Type, TypeVar

from ooo.dyn.style.break_type import BreakType

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase

# endregion Imports

_TBreaks = TypeVar("_TBreaks", bound="Breaks")


class Breaks(StyleBase):
    """
    Paragraph Breaks

    Any properties starting with ``prop_`` set or get current instance values.

    .. seealso::

        - :ref:`help_writer_format_direct_para_text_flow`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, *, type: BreakType | None = None, style: str | None = None, num: int | None = None) -> None:
        """
        Constructor

        Args:
            type (BreakType, optional): Break type.
            style (str, optional): Style to apply to break.
            num (int, optional): Page number to apply to break.

        Returns:
            None:

        Note:
            If argument ``type`` is ``None`` then all other argument are ignored

        See Also:

            - :ref:`help_writer_format_direct_para_text_flow`
        """
        # sourcery skip: merge-nested-ifs
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        # Default for writer is BreakType.NONE
        # BreakType controls the dialog insert checkbox

        # pg_style only applies
        # Dialog position is set by the type: e.g. BreakType.PAGE_AFTER Type is Page and Position is After
        # When BreakType.PAGE_AFTER or COLUMN_AFTER page style is not used.

        if type is None:
            # everything depends on a BreakType
            super().__init__()
            return

        init_vals: Dict[str, Any] = {"BreakType": type}
        if style is not None:
            if type in (BreakType.PAGE_BEFORE, BreakType.COLUMN_BEFORE, BreakType.COLUMN_BOTH, BreakType.PAGE_BOTH):
                # # pg_style is only valid when BreakType.COLUMN_BEFORE or BreakType.PAGE_BEFORE

                # LibreOffice Dev Tools report this property as readonly.
                # api does not.
                # PageStyleName: contains the name of the current page style.

                # init_vals["PageStyleName"] = pg_style

                # PageDescName: If this property is set, it creates a page break before the paragraph
                # it belongs to and assigns the value as the name of the new page style sheet to use.
                init_vals["PageDescName"] = style

        if "PageDescName" in init_vals and num is not None:
            # pg_num is only valid when BreakType.COLUMN_BEFORE or BreakType.PAGE_BEFORE AND
            # page style is set (pg_style)
            init_vals["PageNumberOffset"] = num

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

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies break properties to ``obj``

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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TBreaks], obj: Any) -> _TBreaks: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TBreaks], obj: Any, **kwargs) -> _TBreaks: ...

    @classmethod
    def from_obj(cls: Type[_TBreaks], obj: Any, **kwargs) -> _TBreaks:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Breaks: ``Breaks`` instance that represents ``obj`` break properties.
        """
        # pylint: disable=protected-access
        nu = cls(**kwargs)
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        t = mProps.Props.get(obj, "BreakType", None)
        style = mProps.Props.get(obj, "PageDescName", None)
        num = mProps.Props.get(obj, "PageNumberOffset", None)
        result = cls(type=t, style=style, num=num, **kwargs)
        result.set_update_obj(obj)
        return result

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
    def prop_type(self) -> BreakType | None:
        """Gets break type"""
        return self._get("BreakType")

    @property
    def prop_style(self) -> str | None:
        """Gets Break Style"""
        return self._get("PageDescName")

    @property
    def prop_num(self) -> int | None:
        """Gets Page number to apply to break"""
        return self._get("PageNumberOffset")

    @property
    def default(self: _TBreaks) -> _TBreaks:
        """Gets ``Breaks`` default."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(type=BreakType.NONE, _cattribs=self._get_internal_cattribs())
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
