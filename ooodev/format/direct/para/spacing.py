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
from ...kind.style_kind import StyleKind
from ...style_base import StyleBase


class Spacing(StyleBase):
    """
    Paragraph Spacing

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        above: float | None = None,
        below: float | None = None,
        style_no_space: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            above (float, optional): Determines the top margin of the paragraph (in mm units).
            below (float, optional): Determines the bottom margin of the paragraph (in mm units).
            style_no_space (bool, optional): Do not add space between paragraphs of the same style.
        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        if not above is None:
            init_vals["ParaTopMargin"] = round(above * 100)

        if not below is None:
            init_vals["ParaBottomMargin"] = round(below * 100)

        if not style_no_space is None:
            init_vals["ParaContextMargin"] = style_no_space
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
    def from_obj(obj: object) -> Spacing:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            Spacing: ``Spacing`` instance that represents ``obj`` spacing.
        """
        inst = Spacing()
        if not inst._is_valid_service(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        def set_prop(key: str, indent: Spacing):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                indent._set(key, val)

        set_prop("ParaTopMargin", inst)
        set_prop("ParaBottomMargin", inst)
        set_prop("ParaContextMargin", inst)
        return inst

    # endregion methods

    # region style methods
    def fmt_above(self, value: float | None) -> Spacing:
        """
        Gets a copy of instance with above margin set or removed

        Args:
            value (float | None): Margin value (in mm units).

        Returns:
            Spacing: Indent instance
        """
        cp = self.copy()
        cp.prop_above = value
        return cp

    def fmt_below(self, value: float | None) -> Spacing:
        """
        Gets a copy of instance with below margin set or removed

        Args:
            value (float | None): Margin value (in mm units).

        Returns:
            Spacing: Indent instance
        """
        cp = self.copy()
        cp.prop_below = value
        return cp

    def fmt_style_no_space(self, value: bool | None) -> Spacing:
        """
        Gets a copy of instance with style no spacing set or removed

        Args:
            value (bool | None): Auto value.

        Returns:
            Spacing: Indent instance
        """
        cp = self.copy()
        cp.prop_style_no_space = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def style_no_space(self) -> Spacing:
        """Gets copy of instance with style no spacing set"""
        cp = self.copy()
        cp.prop_style_no_space = True
        return cp

    # endregion Style Properties

    # region properties
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.PARA

    @property
    def prop_above(self) -> float | None:
        """Gets/Sets the top margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaTopMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_above.setter
    def prop_above(self, value: float | None):
        if value is None:
            self._remove("ParaTopMargin")
            return
        self._set("ParaTopMargin", value)

    @property
    def prop_below(self) -> float | None:
        """Gets/Sets the bottom margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaBottomMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_below.setter
    def prop_below(self, value: float | None):
        if value is None:
            self._remove("ParaBottomMargin")
            return
        self._set("ParaBottomMargin", value)

    @property
    def prop_style_no_space(self) -> bool | None:
        """Gets/Sets if no space between paragraphs of the same style"""
        return self._get("ParaContextMargin")

    @prop_style_no_space.setter
    def prop_style_no_space(self, value: bool | None):
        if value is None:
            self._remove("ParaContextMargin")
            return
        self._set("ParaContextMargin", value)

    @static_prop
    def default() -> Spacing:  # type: ignore[misc]
        """Gets ``Spacing`` default. Static Property."""
        if Spacing._DEFAULT is None:
            Spacing._DEFAULT = Spacing(above=0.0, below=0.0, style_no_space=False)
        return Spacing._DEFAULT

    # endregion properties
