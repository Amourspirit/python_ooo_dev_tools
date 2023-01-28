"""
Modele for managing paragraph line numbrs.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase

# from ...events.args.key_val_cancel_args import KeyValCancelArgs


class LineNum(StyleBase):
    """
    Paragraph Line Numbers

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(self, num_start: int = 0) -> None:
        """
        Constructor

        Args:
            num_start (int, optional): Restart paragraph with number.
                If ``0`` then this paragraph is include in line numbering.
                If ``-1`` then this paragraph is excluded in line numbering.
                If greater then zero then this paragraph is included in line numbering and the numbering is restarted with value of ``num_start``.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        init_vals = {}

        if num_start == 0:
            init_vals["ParaLineNumberStartValue"] = 0
            init_vals["ParaLineNumberCount"] = True
        elif num_start < 0:
            init_vals["ParaLineNumberStartValue"] = 0
            init_vals["ParaLineNumberCount"] = False
        else:
            init_vals["ParaLineNumberStartValue"] = num_start
            init_vals["ParaLineNumberCount"] = True

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

    @staticmethod
    def from_obj(obj: object) -> LineNum:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            LineNum: ``LineNum`` instance that represents ``obj`` properties.
        """
        inst = LineNum()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        def set_prop(key: str, o: LineNum):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                o._set(key, val)

        set_prop("ParaLineNumberStartValue", inst)
        set_prop("ParaLineNumberCount", inst)

        return inst

    # endregion methods

    # region Style Methods

    def fmt_num_start(self, value: int) -> LineNum:
        """
        Gets a copy of instance with before list style set or removed

        Args:
            value (int | None): List style value.
                If ``0`` then this paragraph is include in line numbering.
                If ``-1`` then this paragraph is excluded in line numbering.
                If greater then zero then this paragraph is included in line numbering and the numbering is restarted with ``value``.

        Returns:
            LineNum: Line Number instance
        """
        cp = self.copy()
        cp.prop_num_start = value
        return cp

    # endregion Style Methods

    # region Style Properties
    @property
    def restart_numbers(self) -> LineNum:
        """Gets instance with restart numbers set to ``1``"""
        return LineNum(1)

    @property
    def include(self) -> LineNum:
        """Gets instance with include in line numbering set to include."""
        cp = self.copy()
        # zero or higher is already include
        if cp.prop_num_start < 0:
            cp.prop_num_start = 0
        return cp

    @property
    def exclude(self) -> LineNum:
        """Gets instance with include in line numbering set to exclude."""
        return LineNum(-1)

    # endregion Style Properties

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def prop_num_start(self) -> int | None:
        """
        Gets/Sets Restart at this paragraph number.

        If Less then zero then restart numbering at current paragraph is consider to be ``False``;
        Otherewise; restart numbering is considered to be ``True``.
        """
        return self._get("ParaLineNumberStartValue")

    @prop_num_start.setter
    def prop_num_start(self, value: int) -> None:
        if value == 0:
            self._set("ParaLineNumberStartValue", 0)
            self._set("ParaLineNumberCount", True)
            return
        if value < 0:
            self._set("ParaLineNumberStartValue", 0)
            self._set("ParaLineNumberCount", False)
        self._set("ParaLineNumberStartValue", value)
        self._set("ParaLineNumberCount", True)

    @static_prop
    def default() -> LineNum:  # type: ignore[misc]
        """Gets ``LineNum`` default. Static Property."""
        if LineNum._DEFAULT is None:
            LineNum._DEFAULT = LineNum(0)
        return LineNum._DEFAULT

    # endregion properties
