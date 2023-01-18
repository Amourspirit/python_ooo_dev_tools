"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload

from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import lo as mLo
from ....utils import props as mProps
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase


class Indent(StyleBase):
    """
    Paragraph Indent

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        before: float | None = None,
        after: float | None = None,
        first: float | None = None,
        auto: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            before (float, optional): Determines the left margin of the paragraph (in mm units).
            after (float, optional): Determines the right margin of the paragraph (in mm units).
            first (float, optional): specifies the indent for the first line (in mm units).
            auto (bool, optional): Determines if the first line should be indented automatically.
        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if not before is None:
            init_vals["ParaLeftMargin"] = round(before * 100)

        if not after is None:
            init_vals["ParaRightMargin"] = round(after * 100)

        if not first is None:
            init_vals["ParaFirstLineIndent"] = round(first * 100)

        if not auto is None:
            init_vals["ParaIsAutoFirstLineIndent"] = auto
        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies writing mode to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    @staticmethod
    def from_obj(obj: object) -> Indent:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            WritingMode: ``Indent`` instance that represents ``obj`` writing mode.
        """
        inst = Indent()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        def set_prop(key: str, indent: Indent):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                indent._set(key, val)

        set_prop("ParaLeftMargin", inst)
        set_prop("ParaRightMargin", inst)
        set_prop("ParaFirstLineIndent", inst)
        set_prop("ParaIsAutoFirstLineIndent", inst)
        return inst

    # endregion methods

    # region style methods
    def fmt_before(self, value: float | None) -> Indent:
        """
        Gets a copy of instance with before margin set or removed

        Args:
            value (float | None): Margin value (in mm units).

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_before = value
        return cp

    def fmt_after(self, value: float | None) -> Indent:
        """
        Gets a copy of instance with after margin set or removed

        Args:
            value (float | None): Margin value (in mm units).

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_after = value
        return cp

    def fmt_first(self, value: float | None) -> Indent:
        """
        Gets a copy of instance with first indent margin set or removed

        Args:
            value (float | None): Margin value (in mm units).

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_after = value
        return cp

    def fmt_auto(self, value: bool | None) -> Indent:
        """
        Gets a copy of instance with auto set or removed

        Args:
            value (bool | None): Auto value.

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_auto = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def auto(self) -> Indent:
        """Gets copy of instance with auto set"""
        cp = self.copy()
        cp.prop_auto = True
        return cp

    # endregion Style Properties

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def prop_before(self) -> float | None:
        """Gets/Sets the left margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaLeftMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_before.setter
    def prop_before(self, value: float | None):
        if value is None:
            self._remove("ParaLeftMargin")
            return
        self._set("ParaLeftMargin", value)

    @property
    def prop_after(self) -> float | None:
        """Gets/Sets the right margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaRightMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_after.setter
    def prop_after(self, value: float | None):
        if value is None:
            self._remove("ParaRightMargin")
            return
        self._set("ParaRightMargin", value)

    @property
    def prop_first(self) -> float | None:
        """Gets/Sets the indent for the first line (in mm units)."""
        pv = cast(int, self._get("ParaFirstLineIndent"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_first.setter
    def prop_first(self, value: float | None):
        if value is None:
            self._remove("ParaFirstLineIndent")
            return
        self._set("ParaFirstLineIndent", value)

    @property
    def prop_auto(self) -> bool | None:
        """Gets/Sets if the first line should be indented automatically"""
        return self._get("ParaIsAutoFirstLineIndent")

    @prop_auto.setter
    def prop_auto(self, value: bool | None):
        if value is None:
            self._remove("ParaIsAutoFirstLineIndent")
            return
        self._set("ParaIsAutoFirstLineIndent", value)

    @static_prop
    def default() -> Indent:  # type: ignore[misc]
        """Gets ``Indent`` default. Static Property."""
        if Indent._DEFAULT is None:
            Indent._DEFAULT = Indent(before=0.0, after=0.0, first=0.0, auto=False)
        return Indent._DEFAULT

    # endregion properties
