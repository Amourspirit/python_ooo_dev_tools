from __future__ import annotations
from typing import Any, Tuple, overload

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.meta.static_prop import static_prop
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleName
from ooodev.utils import props as mProps
from ooodev.exceptions import ex as mEx
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind


class Page(StyleName):
    """
    Page Style.

    .. seealso::

        - :ref:`help_calc_format_style_page`

    .. versionadded:: 0.9.0
    """

    def __init__(self, name: CalcStylePageKind | str = "") -> None:
        """
        Constructor

        Args:
            name (CalcStylePageKind | str, optional): Style name. Defaults to ``Default``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_page`
        """
        if name == "":
            name = Page.default.prop_name
        super().__init__(name=name)

    # region Overrides

    # region Copy()
    @overload
    def copy(self) -> Page:
        ...

    @overload
    def copy(self, **kwargs) -> Page:
        ...

    def copy(self, **kwargs) -> Page:
        """Gets a copy of instance as a new instance"""
        return Page(name=self.prop_name)

    # endregion Copy()
    def _get_family_style_name(self) -> str:
        return "PageStyles"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.sheet.Spreadsheet",)
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._style_property_name
        except AttributeError:
            self._style_property_name = "PageStyle"
        return self._style_property_name

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs):
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        if event_args.value == "":
            event_args.value = Page.default.prop_name
        super().on_property_setting(source, event_args)

    # endregion Overrides

    # region Static Methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> Page:
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Page:
        ...

    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> Page:
        """
        Gets instance from object

        Args:
            obj (Any): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Page: ``Page`` instance that represents ``obj`` style.
        """
        # Write Property for page is PageDescName and read property is PageStyleName.
        inst = Page(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        if name := mProps.Props.get(obj, inst._get_property_name(), ""):
            inst.prop_name = name
        return inst

    # endregion from_obj()

    # endregion Static Methods

    # region Style Properties
    @property
    def report(self) -> Page:
        """Style Report"""
        return Page(CalcStylePageKind.REPORT)

    # endregion Style Properties

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE | FormatKind.PAGE | FormatKind.STATIC
        return self._format_kind_prop

    @static_prop
    def default() -> Page:  # type: ignore[misc]
        """Gets Page default style. Static Property."""
        try:
            return Page._DEFAULT_PAGE  # type: ignore[return-value]
        except AttributeError:
            Page._DEFAULT_PAGE = Page(name=CalcStylePageKind.DEFAULT)  # type: ignore[assignment]
            Page._DEFAULT_PAGE._is_default_inst = True  # type: ignore[assignment]
        return Page._DEFAULT_PAGE  # type: ignore[return-value]

    # endregion Properties
