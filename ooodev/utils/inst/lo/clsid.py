from __future__ import annotations
import uno
from enum import Enum

# region CLSIDs for Office documents
# defined in https://github.com/LibreOffice/core/blob/master/officecfg/registry/data/org/openoffice/Office/Embedding.xcu
# https://opengrok.libreoffice.org/xref/core/officecfg/registry/data/org/openoffice/Office/Embedding.xcu


class CLSID(str, Enum):
    """CLSID for Office documents"""

    # in lower case by design.
    WRITER = "8bc6b165-b1b2-4edd-aa47-dae2ee689dd6"
    CALC = "47bbb4cb-ce4c-4e80-a591-42d9ae74950f"
    DRAW = "4bab8970-8a3b-45b3-991c-cbeeac6bd5e3"
    IMPRESS = "9176e48a-637a-4d1f-803b-99d9bfac1047"
    MATH = "078b7aba-54fc-457f-8551-6147e776a997"
    CHART = "12dcae26-281f-416f-a234-c3086127382e"

    def __str__(self) -> str:
        return self.value


# unsure about these:
#
# chart2 "80243D39-6741-46C5-926E-069164FF87BB"
#       service: com.sun.star.chart2.ChartDocument

#  applet "970B1E81-CF2D-11CF-89CA-008029E4B0B1"
#       service: com.sun.star.comp.sfx2.AppletObject

#  plug-in "4CAA7761-6B8B-11CF-89CA-008029E4B0B1"
#        service: com.sun.star.comp.sfx2.PluginObject

#  frame "1A8A6701-DE58-11CF-89CA-008029E4B0B1"
#        service: com.sun.star.comp.sfx2.IFrameObject

#  XML report chart "D7896D52-B7AF-4820-9DFE-D404D015960F"
#        service: com.sun.star.report.ReportDefinition

# endregion CLSIDs for Office documents
