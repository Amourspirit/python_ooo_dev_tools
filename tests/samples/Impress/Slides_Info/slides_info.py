from __future__ import annotations

import uno
from com.sun.star.lang import XComponent

from ooodev.office.draw import Draw, DrawingLayerKind
from ooodev.utils.file_io import FileIO
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.type_var import PathOrStr


class SlidesInfo:
    def __init__(self, fnm: PathOrStr) -> None:
        FileIO.is_exist_file(fnm=fnm, raise_err=True)
        self._fnm = FileIO.get_absolute_path(fnm)

    def main(self) -> None:
        with Lo.Loader(Lo.ConnectPipe(headless=True)) as loader:

            doc = Lo.open_doc(fnm=self._fnm, loader=loader)

            if not Draw.is_shapes_based(doc):
                Lo.print("-- not a drawing or slides presentation")
                return

            print()
            print(f"No. of slides: {Draw.get_slides_count(doc)}")
            print()

            # Access the first page
            slide = Draw.get_slide(doc=doc, idx=0)

            slide_size = Draw.get_slide_size(slide)
            print(f"Size of slide (mm)({slide_size.Width}, {slide_size.Height})")
            print()

            self._report_layers(doc)
            self._report_styles(doc)
            Lo.close_doc(doc)

    def _report_layers(self, doc: XComponent) -> None:
        lm = Draw.get_layer_manager(doc)
        for i in range(lm.getCount()):
            try:
                Props.show_obj_props(f" Layer {i}", lm.getByIndex(i))
            except:
                pass
        layer = Draw.get_layer(doc=doc, layer_name=DrawingLayerKind.BACK_GROUND_OBJECTS)
        Props.show_obj_props("Background Object Props", layer)

    def _report_styles(self, doc: XComponent) -> None:
        style_names = Info.get_style_family_names(doc)
        print("Style Families in this document:")
        Lo.print_names(style_names)
        # usually: "Default"  "cell"  "graphics"  "table"
        # Default is the name of the default Master Page template inside Office

        for name in style_names:
            con_names = Info.get_style_names(doc=doc, family_style_name=name)
            print(f'Styles in the "{name}" style family:')
            Lo.print_names(con_names)
