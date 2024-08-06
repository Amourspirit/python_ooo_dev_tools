"""
Module for managing paragraph hyphenation.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import Any, Tuple, overload, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase

# endregion Imports

_THyphenation = TypeVar("_THyphenation", bound="Hyphenation")


class Hyphenation(StyleBase):
    """
    Paragraph Hyphenation

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_text_flow`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        auto: bool | None = None,
        no_caps: bool | None = None,
        start_chars: int | None = None,
        end_chars: int | None = None,
        max: int | None = None,
    ) -> None:
        """
        Constructor

        Args:
            auto (bool, optional): Hyphenate automatically.
            no_caps (bool, optional): Don't hyphenate word in caps.
            start_chars (int, optional): Characters at line begin.
            end_chars (int, optional): characters at line end.
            max (int, optional): Maximum consecutive hyphenated lines.

        Returns:
            None:

        Note:
            If argument ``auto`` is ``False`` then all other argument have no effect.

        See Also:

            - :ref:`help_writer_format_direct_para_text_flow`
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}
        if auto is not None:
            init_vals["ParaIsHyphenation"] = auto

        if start_chars is not None:
            init_vals["ParaHyphenationMaxLeadingChars"] = start_chars

        if end_chars is not None:
            init_vals["ParaHyphenationMaxTrailingChars"] = end_chars

        if no_caps is not None:
            init_vals["ParaHyphenationNoCaps"] = no_caps

        if max is not None:
            init_vals["ParaHyphenationMaxHyphens"] = max

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
        Applies hyphenation properties to ``obj``

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
    def from_obj(cls: Type[_THyphenation], obj: Any) -> _THyphenation: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_THyphenation], obj: Any, **kwargs) -> _THyphenation: ...

    @classmethod
    def from_obj(cls: Type[_THyphenation], obj: Any, **kwargs) -> _THyphenation:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Hyphenation: ``Hyphenation`` instance that represents ``obj`` hyphenation properties.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, indent: Hyphenation):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                indent._set(key, val)

        set_prop("ParaIsHyphenation", inst)
        set_prop("ParaHyphenationMaxLeadingChars", inst)
        set_prop("ParaHyphenationMaxTrailingChars", inst)
        set_prop("ParaHyphenationNoCaps", inst)
        set_prop("ParaHyphenationMaxHyphens", inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region style methods
    def fmt_auto(self: _THyphenation, value: bool | None) -> _THyphenation:
        """
        Gets copy of instance with auto set or removed

        Args:
            value (bool | None): auto value

        Returns:
            Hyphenation: ``Hyphenation`` instance
        """
        cp = self.copy()
        cp.prop_auto = value
        return cp

    def fmt_no_caps(self: _THyphenation, value: bool | None) -> _THyphenation:
        """
        Gets copy of instance with no caps set or removed

        Args:
            value (bool | None): no caps value

        Returns:
            Hyphenation: ``Hyphenation`` instance
        """
        cp = self.copy()
        cp.prop_no_caps = value
        return cp

    def fmt_start_chars(self: _THyphenation, value: int | None) -> _THyphenation:
        """
        Gets copy of instance with start chars set or removed

        Args:
            value (bool | None): start chars value

        Returns:
            Hyphenation: ``Hyphenation`` instance
        """
        cp = self.copy()
        cp.prop_start_chars = value
        return cp

    def fmt_end_chars(self: _THyphenation, value: int | None) -> _THyphenation:
        """
        Gets copy of instance with end chars set or removed

        Args:
            value (bool | None): end chars value

        Returns:
            Hyphenation: ``Hyphenation`` instance
        """
        cp = self.copy()
        cp.prop_end_chars = value
        return cp

    def fmt_max_chars(self: _THyphenation, value: int | None) -> _THyphenation:
        """
        Gets copy of instance with max set or removed

        Args:
            value (bool | None): max value

        Returns:
            Hyphenation: ``Hyphenation`` instance
        """
        cp = self.copy()
        cp.prop_max = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def auto(self: _THyphenation) -> _THyphenation:
        """
        Gets instance with Hyphenate automatically set to ``True``.
        """
        cp = self.copy()
        cp.prop_auto = True
        return cp

    @property
    def no_caps(self: _THyphenation) -> _THyphenation:
        """
        Gets instance with no caps set to ``True``.
        """
        cp = self.copy()
        cp.prop_no_caps = True
        return cp

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
    def prop_auto(self) -> bool | None:
        """Gets/Sets Hyphenate automatically."""
        return self._get("ParaIsHyphenation")

    @prop_auto.setter
    def prop_auto(self, value: bool | None):
        if value is None:
            self._remove("ParaIsHyphenation")
            return
        self._set("ParaIsHyphenation", value)

    @property
    def prop_no_caps(self) -> bool | None:
        """Gets/Sets if hyphenate word in caps."""
        return self._get("ParaHyphenationNoCaps")

    @prop_no_caps.setter
    def prop_no_caps(self, value: bool | None):
        if value is None:
            self._remove("ParaHyphenationNoCaps")
            return
        self._set("ParaHyphenationNoCaps", value)

    @property
    def prop_start_chars(self) -> int | None:
        """Gets/Sets number of characters at line begin."""
        return self._get("ParaHyphenationMaxLeadingChars")

    @prop_start_chars.setter
    def prop_start_chars(self, value: int | None):
        if value is None:
            self._remove("ParaHyphenationMaxLeadingChars")
            return
        self._set("ParaHyphenationMaxLeadingChars", value)

    @property
    def prop_end_chars(self) -> int | None:
        """Gets/Sets number of characters at line end."""
        return self._get("ParaHyphenationMaxTrailingChars")

    @prop_end_chars.setter
    def prop_end_chars(self, value: int | None):
        if value is None:
            self._remove("ParaHyphenationMaxTrailingChars")
            return
        self._set("ParaHyphenationMaxTrailingChars", value)

    @property
    def prop_max(self) -> int | None:
        """Gets/Sets maximum consecutive hyphenated lines."""
        return self._get("ParaHyphenationMaxHyphens")

    @prop_max.setter
    def prop_max(self, value: int | None):
        if value is None:
            self._remove("ParaHyphenationMaxHyphens")
            return
        self._set("ParaHyphenationMaxHyphens", value)

    @property
    def default(self: _THyphenation) -> _THyphenation:
        """Gets ``Hyphenation`` default. Static Property."""
        # pylint: disable=unexpected-keyword-arg
        # pylint: disable=protected-access
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(
                auto=False, no_caps=False, start_chars=2, end_chars=2, max=0, _cattribs=self._get_internal_cattribs()
            )
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
