# region Imports
from __future__ import annotations
from typing import Any, Tuple
import uno
from com.sun.star.beans import XPropertySet

from ooodev.exceptions import ex as mEx
from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo

# endregion Imports


class CellStyleBase(StyleBase):
    """
    Cell Style Base

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
        **kwargs,
    ) -> None:
        """
        Constructor

        Args:
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:
        """
        # sourcery skip: remove-unnecessary-cast
        self._style_name = str(style_name)
        self._style_family_name = str(style_family)
        super().__init__(**kwargs)

    # endregion Init

    # region internal methods

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CellStyle",
                "com.sun.star.style.PageStyle",
            )
        return self._supported_services_values

    def _is_valid_doc(self, obj: Any) -> bool:
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.CALC)

    # endregion internal methods

    # region Methods

    def get_style_props(self, doc: object) -> XPropertySet:
        """
        Gets the Style Properties

        Args:
            doc (object): UNO Document Object.

        Raised:
            NotSupportedDocumentError: If document is not supported.

        Returns:
            XPropertySet: Styles properties property set.
        """
        if not self._is_valid_doc(doc):
            raise mEx.NotSupportedDocumentError
        return mInfo.Info.get_style_props(doc, self.prop_style_family_name, self.prop_style_name)

    # endregion Methods

    # region Overrides

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO Calc Document

        Returns:
            None:
        """

        if not self._is_valid_doc(obj):
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Not a Valid Document. Unable to set Style Property")
            return
        p = self.get_style_props(obj)
        super().apply(p, **kwargs)

    # endregion Overrides

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PAGE | FormatKind.STYLE
        return self._format_kind_prop

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | CalcStylePageKind):
        self._style_name = str(value)

    @property
    def prop_style_family_name(self) -> str:
        """Gets/Set Style Family Name"""
        return self._style_family_name

    @prop_style_family_name.setter
    def prop_style_family_name(self, value: str):
        self._style_family_name = value

    # endregion Properties
