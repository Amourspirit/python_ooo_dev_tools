import uno
from ooo.dyn.table.border_line_style import BorderLineStyleEnum as BorderLineStyleEnum
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation
from ....style.para import StyleParaKind as StyleParaKind
from .....direct.structs.side import Side as Side, SideFlags as SideFlags, LineSize as LineSize
from .....direct.para.border.padding import Padding as DirectPadding
from .....direct.para.border.shadow import Shadow as DirectShadow
from .....modify.para.border.padding import Padding as Padding
from .....modify.para.border.sides import Sides as Sides
from .....modify.para.border.shadow import Shadow as Shadow
from .....modify.para.border.borders import Borders as Borders
