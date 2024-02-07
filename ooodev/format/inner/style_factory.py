from __future__ import annotations
from typing import TYPE_CHECKING, Type

if TYPE_CHECKING:
    from .proto.font_effects_t import FontEffectsT


def font_effects_factory(name: str) -> Type[FontEffectsT]:
    if name == "ooodev.write.char":
        from ooodev.format.inner.direct.write.char.font.font_effects import FontEffects

        return FontEffects
    if name == "ooodev.chart2.axis":
        from ooodev.format.inner.direct.chart2.axis.font.font_effects import FontEffects

        return FontEffects

    if name == "ooodev.chart2.title":
        from ooodev.format.chart2.direct.title.font import FontEffects

        return FontEffects

    if name == "ooodev.chart2.series.data_labels":
        from ooodev.format.chart2.direct.series.data_labels.font import FontEffects

        return FontEffects

    raise ValueError(f"Invalid name: {name}")
