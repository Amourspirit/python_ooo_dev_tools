from __future__ import annotations
import uno
from com.sun.star.awt import XBitmap

from ooodev.format.inner.common.props.area_pattern_props import AreaPatternProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.write.page.header.area.pattern import Pattern as HeaderPattern
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind


class Pattern(HeaderPattern):
    """
    Page Footer Pattern

    .. seealso::

        - :ref:`help_writer_format_modify_page_footer_area`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this
                property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_footer_area`
        """
        super().__init__(
            bitmap=bitmap,
            name=name,
            tile=tile,
            stretch=stretch,
            auto_name=auto_name,
            style_name=style_name,
            style_family=style_family,
        )

    # region Internal Methods
    def _get_inner_props(self) -> AreaPatternProps:
        return AreaPatternProps(
            style="FooterFillStyle",
            name="FooterFillBitmapName",
            tile="FooterFillBitmapTile",
            stretch="FooterFillBitmapStretch",
            bitmap="FooterFillBitmap",
        )

    # endregion Internal Methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion properties
