from __future__ import annotations
from typing import Tuple, overload
import uno
from com.sun.star.awt import XBitmap
from com.sun.star.chart2 import XChartDocument
from com.sun.star.lang import XMultiServiceFactory

from ooo.dyn.drawing.fill_style import FillStyle as FillStyle

from ooodev.format.inner.common.props.area_pattern_props import AreaPatternProps
from ooodev.format.inner.direct.write.fill.area.pattern import Pattern as FillPattern
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset import preset_pattern as mPattern
from ooodev.format.inner.preset.preset_pattern import PresetPatternKind as PresetPatternKind
from ooodev.meta.deleted_attrib import DeletedAttrib
from ooodev.utils import lo as mLo


class Pattern(FillPattern):
    """
    Class for Chart Area Fill Pattern.

    .. versionadded:: 0.9.4
    """

    prop_tile = DeletedAttrib()
    prop_stretch = DeletedAttrib()

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
    ) -> None:
        """
        _summary_

        Args:
            chart_doc (XChartDocument): Chart document.
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
        """
        self._chart_doc = chart_doc
        super().__init__(bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name)

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.PageBackground",
                "com.sun.star.chart2.Title",
                "com.sun.star.chart2.DataSeries",
            )
        return self._supported_services_values

    def _container_get_msf(self) -> XMultiServiceFactory | None:
        if self._chart_doc is not None:
            chart_doc_ms_factory = mLo.Lo.qi(XMultiServiceFactory, self._chart_doc)
            return chart_doc_ms_factory
        return None

    # region copy()
    @overload
    def copy(self) -> Pattern:
        ...

    @overload
    def copy(self, **kwargs) -> Pattern:
        ...

    def copy(self, **kwargs) -> Pattern:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._chart_doc = self._chart_doc
        return cp

    # endregion copy()
    # endregion overrides

    # region Static Methods
    # region from_preset()
    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetPatternKind) -> Pattern:
        ...

    @overload
    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetPatternKind, **kwargs) -> Pattern:
        ...

    @classmethod
    def from_preset(cls, chart_doc: XChartDocument, preset: PresetPatternKind, **kwargs) -> Pattern:
        """
        Gets an instance from a preset.

        Args:
            preset (~.preset.preset_pattern.PresetPatternKind): Preset.

        Returns:
            Pattern: ``Pattern`` instance from preset.
        """
        name = str(preset)
        nu = cls(chart_doc=chart_doc, **kwargs)

        nc = nu._container_get_inst()
        bmap = nu._container_get_value(name, nc)
        if bmap is None:
            bmap = mPattern.get_prest_bitmap(preset)
        return cls(chart_doc=chart_doc, bitmap=bmap, name=name, tile=True, stretch=False, auto_name=False, **kwargs)

    # endregion from_preset()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object) -> Pattern:
        ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object, **kwargs) -> Pattern:
        ...

    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: object, **kwargs) -> Pattern:
        """
        Gets instance from object

        Args:
            chart_doc (XChartDocument): Chart document.
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Pattern: ``Pattern`` instance that represents ``obj`` fill pattern.
        """
        return super().from_obj(obj=obj, chart_doc=chart_doc, **kwargs)

    # endregion from_obj()
    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.FILL
        return self._format_kind_prop

    @property
    def _props(self) -> AreaPatternProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaPatternProps(
                style="FillStyle",
                name="FillBitmapName",
                tile="",
                stretch="",
                bitmap="",
            )
        return self._props_internal_attributes

    # endregion Properties
