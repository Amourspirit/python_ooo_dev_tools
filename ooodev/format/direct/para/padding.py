"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload

from ....exceptions import ex as mEx
from ....meta.static_prop import static_prop
from ....utils import info as mInfo
from ....utils import lo as mLo
from ....utils import props as mProps
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase


class Padding(StyleBase):
    """
    Paragraph Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        padding_all: float | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Paragraph left padding (in mm units).
            right (float, optional): Paragraph right padding (in mm units).
            top (float, optional): Paragraph top padding (in mm units).
            bottom (float, optional): Paragraph bottom padding (in mm units).
            padding_all (float, optional): Paragraph left, right, top, bottom padding (in mm units). If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.

        Raises:
            ValueError: If any argument value is less than zero.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        def validate(val: float | None) -> None:
            if val is not None:
                if val < 0.0:
                    raise ValueError("padding values must be positive values")

        def set_val(key, value) -> None:
            nonlocal init_vals
            if not value is None:
                init_vals[key] = round(value * 100)

        validate(left)
        validate(right)
        validate(top)
        validate(bottom)
        validate(padding_all)
        para_attrs = ("ParaLeftMargin", "ParaRightMargin", "ParaTopMargin", "ParaBottomMargin")
        if padding_all is None:
            for key, value in zip(para_attrs, (left, right, top, bottom)):
                set_val(key, value)
        else:
            for key in para_attrs:
                set_val(key, padding_all)

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
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")
        return None

    # endregion apply()

    @staticmethod
    def from_obj(obj: object) -> Padding:
        """
        Gets Padding instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            Padding: Padding that represents ``obj`` padding.
        """
        if not mInfo.Info.support_service(obj, "com.sun.star.style.ParagraphProperties"):
            raise mEx.NotSupportedServiceError("com.sun.star.style.ParagraphProperties")
        inst = Padding()
        if inst._is_valid_obj(obj):
            inst._set("ParaLeftMargin", int(mProps.Props.get(obj, "ParaLeftMargin")))
            inst._set("ParaRightMargin", int(mProps.Props.get(obj, "ParaRightMargin")))
            inst._set("ParaTopMargin", int(mProps.Props.get(obj, "ParaTopMargin")))
            inst._set("ParaBottomMargin", int(mProps.Props.get(obj, "ParaBottomMargin")))
        else:
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])
        return inst

    # endregion methods

    # region style methods
    def fmt_padding_all(self, value: float | None) -> Padding:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_top = value
        cp.prop_bottom = value
        cp.prop_left = value
        cp.prop_right = value
        return cp

    def fmt_top(self, value: float | None) -> Padding:
        """
        Gets a copy of instance with top side set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def fmt_bottom(self, value: float | None) -> Padding:
        """
        Gets a copy of instance with bottom side set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def fmt_left(self, value: float | None) -> Padding:
        """
        Gets a copy of instance with left side set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def fmt_right(self, value: float | None) -> Padding:
        """
        Gets a copy of instance with right side set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_right = value
        return cp

    # endregion style methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def prop_left(self) -> float | None:
        """Gets/Sets paragraph left padding (in mm units)."""
        pv = cast(int, self._get("ParaLeftMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_left.setter
    def prop_left(self, value: float | None):
        if value is None:
            self._remove("ParaLeftMargin")
            return
        self._set("ParaLeftMargin", round(value * 100))

    @property
    def prop_right(self) -> float | None:
        """Gets/Sets paragraph right padding (in mm units)."""
        pv = cast(int, self._get("ParaRightMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_right.setter
    def prop_right(self, value: float | None):
        if value is None:
            self._remove("ParaRightMargin")
            return
        self._set("ParaRightMargin", round(value * 100))

    @property
    def prop_top(self) -> float | None:
        """Gets/Sets paragraph top padding (in mm units)."""
        pv = cast(int, self._get("ParaTopMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_top.setter
    def prop_top(self, value: float | None):
        if value is None:
            self._remove("ParaTopMargin")
            return
        self._set("ParaTopMargin", round(value * 100))

    @property
    def prop_bottom(self) -> float | None:
        """Gets/Sets paragraph bottom padding (in mm units)."""
        pv = cast(int, self._get("ParaBottomMargin"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_bottom.setter
    def prop_bottom(self, value: float | None):
        if value is None:
            self._remove("ParaBottomMargin")
            return
        self._set("ParaBottomMargin", round(value * 100))

    @static_prop
    def default() -> Padding:  # type: ignore[misc]
        """Gets Padding default. Static Property."""
        if Padding._DEFAULT is None:
            Padding._DEFAULT = Padding(padding_all=0.35)
        return Padding._DEFAULT

    # endregion properties
