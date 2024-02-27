# region Import
from __future__ import annotations
from typing import Tuple, cast, Type, TypeVar
import uno

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.common.abstract.abstract_hf import AbstractHF
from ooodev.format.inner.common.props.hf_props import HfProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.meta.deleted_attrib import DeletedAttrib
from ooodev.meta.disabled_method import DisabledMethod

# endregion Import

_THeader = TypeVar(name="_THeader", bound="Header")


class InnerStyle(AbstractHF):
    """
    Page Header/Footer Inner Settings

    .. versionadded:: 0.9.0
    """

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")
        return self._supported_services_values

    # endregion Overrides

    # region Properties

    @property
    def _props(self) -> HfProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = HfProps(
                on="HeaderIsOn",
                shared="HeaderIsShared",
                shared_first="FirstPageHeaderIsShared",
                margin_left="HeaderLeftMargin",
                margin_right="HeaderRightMargin",
                spacing="HeaderBodyDistance",
                spacing_dyn="HeaderDynamic",
                height="HeaderHeight",
                height_auto="HeaderIsDynamicHeight",
            )
        return self._props_internal_attributes

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.HEADER
        return self._format_kind_prop

    # endregion Properties


class Header(CellStyleBaseMulti):
    """
    Page Header Settings

    .. seealso::

        - :ref:`help_calc_format_modify_page_header_header`

    .. versionadded:: 0.9.0
    """

    space_dyn = DeletedAttrib()
    fmt_spacing_dyn = DisabledMethod()

    def __init__(
        self,
        *,
        on: bool | None = None,
        shared: bool | None = None,
        shared_first: bool | None = None,
        margin_left: float | None = None,
        margin_right: float | None = None,
        spacing: float | None = None,
        height: float | None = None,
        height_auto: bool | None = None,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            on (bool | None, optional): Specifics if Footer is on.
            shared (bool | None, optional): Specifies if same contents left and right.
            shared_first (bool | None, optional): Specifies if same contents on first page.
            margin_left (float | None, optional): Specifies Left Margin in ``mm`` units.
            margin_right (float | None, optional): Specifies Right Margin in ``mm`` units.
            spacing (float | None, optional): Specifies Spacing in ``mm`` units.
            height (float | None, optional): Specifies Height in ``mm`` units.
            height_auto (bool | None, optional): Specifies if auto-fit height is used.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_header_header`
        """

        direct = InnerStyle(
            on=on,
            shared=shared,
            shared_first=shared_first,
            margin_left=margin_left,
            margin_right=margin_right,
            spacing=spacing,
            height=height,
            height_auto=height_auto,
            _cattribs=self._get_inner_cattribs(),
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    # region Internal Methods
    def _get_inner_props(self) -> HfProps:
        return HfProps(
            on="HeaderIsOn",
            shared="HeaderIsShared",
            shared_first="FirstPageHeaderIsShared",
            margin_left="HeaderLeftMargin",
            margin_right="HeaderRightMargin",
            spacing="HeaderBodyDistance",
            spacing_dyn="HeaderDynamic",
            height="HeaderHeight",
            height_auto="HeaderIsDynamicHeight",
        )

    def _get_inner_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_format_kind_prop": self.prop_format_kind,
            "_props_internal_attributes": self._get_inner_props(),
        }

    # endregion Internal Methods

    # region Static Methods
    @classmethod
    def from_style(
        cls: Type[_THeader],
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> _THeader:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Header: ``Header`` or ``Footer`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerStyle.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct)
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.HEADER
        return self._format_kind_prop

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | CalcStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerStyle:
        """Gets/Sets Inner Footer instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerStyle, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerStyle) -> None:
        if not isinstance(value, InnerStyle):
            raise TypeError(f'Expected type of InnerStyle, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value)

    # endregion Properties
