"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload

from ...exceptions import ex as mEx
from ...meta.static_prop import static_prop
from ...utils import info as mInfo
from ...utils import lo as mLo
from ...utils import props as mProps
from ..kind.style_kind import StyleKind
from ..style_base import StyleBase


class Indent(StyleBase):
    """
    Paragraph Alignment

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``style_`` can be used to chain together Padding properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        before_text: float | None = None,
        after_text: float | None = None,
        first_indent: float | None = None,
        auto: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            before_text (float, optional): Determines the left margin of the paragraph (in mm units).
            after_text (float, optional): Determines the right margin of the paragraph (in mm units).
            first_indent (float, optional): specifies the indent for the first line (in mm units).
            auto (bool, optional): Determines if the first line should be indented automatically.
        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if not before_text is None:
            init_vals["ParaLeftMargin"] = round(before_text * 100)

        if not after_text is None:
            init_vals["ParaRightMargin"] = round(after_text * 100)

        if not first_indent is None:
            init_vals["ParaFirstLineIndent"] = round(first_indent * 100)

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

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies writing mode to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphPropertiesComplex`` service.

        Returns:
            None:
        """
        try:
            super().apply_style(obj)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @staticmethod
    def from_obj(obj: object) -> Indent:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphPropertiesComplex`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphPropertiesComplex`` service.

        Returns:
            WritingMode: ``WritingMode`` instance that represents ``obj`` writing mode.
        """
        inst = Indent()
        if not inst._is_valid_service(obj):
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
    def style_before_text(self, value: float | None) -> Indent:
        """
        Gets a copy of instance with before text set or removed

        Args:
            value (float | None): Margin value (in mm units).

        Returns:
            Padding: Indent instance
        """
        cp = self.copy()
        cp.prop_before_text = value
        return cp

    def style_after_text(self, value: float | None) -> Indent:
        """
        Gets a copy of instance with after text set or removed

        Args:
            value (float | None): Margin value (in mm units).

        Returns:
            Padding: Indent instance
        """
        cp = self.copy()
        cp.prop_after_text = value
        return cp

    def style_first_indent(self, value: float | None) -> Indent:
        """
        Gets a copy of instance with first indent set or removed

        Args:
            value (float | None): Margin value (in mm units).

        Returns:
            Padding: Indent instance
        """
        cp = self.copy()
        cp.prop_after_text = value
        return cp

    def style_auto(self, value: bool | None) -> Indent:
        """
        Gets a copy of instance with auto set or removed

        Args:
            value (bool | None): Auto value.

        Returns:
            Padding: Indent instance
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
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.PARA

    @property
    def prop_before_text(self) -> float | None:
        """Gets/Sets the left margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaLeftMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_before_text.setter
    def prop_before_text(self, value: float | None):
        if value is None:
            self._remove("ParaLeftMargin")
            return
        self._set("ParaLeftMargin", value)

    @property
    def prop_after_text(self) -> float | None:
        """Gets/Sets the right margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaRightMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_after_text.setter
    def prop_after_text(self, value: float | None):
        if value is None:
            self._remove("ParaRightMargin")
            return
        self._set("ParaRightMargin", value)

    @property
    def prop_first_indent(self) -> float | None:
        """Gets/Sets the indent for the first line (in mm units)."""
        pv = cast(int, self._get("ParaFirstLineIndent"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_first_indent.setter
    def prop_first_indent(self, value: float | None):
        if value is None:
            self._remove("ParaFirstLineIndent")
            return
        self._set("ParaFirstLineIndent", value)

    @property
    def prop_auto(self) -> bool | None:
        """Gets/Sets  if the first line should be indented automatically"""
        return self._get("ParaIsAutoFirstLineIndent")

    @prop_auto.setter
    def prop_auto(self, value: bool | None):
        if value is None:
            self._remove("ParaIsAutoFirstLineIndent")
            return
        self._set("ParaIsAutoFirstLineIndent", value)

    @static_prop
    def default(cls) -> Indent:
        """Gets ``WritingMode`` default. Static Property."""
        if cls._DEFAULT is None:
            cls._DEFAULT = Indent(before_text=0.0, after_text=0.0, first_indent=0.0, auto=False)
        return cls._DEFAULT

    # endregion properties
