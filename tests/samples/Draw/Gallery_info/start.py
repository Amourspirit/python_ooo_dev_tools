from __future__ import annotations
import sys
import uno
from ooodev.utils.lo import Lo
from ooodev.utils.gallery import Gallery, SearchMatchKind, GalleryKind, SearchByKind
from ooodev.exceptions.ex import GalleryNotFoundError
from ooodev.utils.info import Info


class GalleryInfo:
    @staticmethod
    def main() -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:

            # list all the gallery themes (i.e. the sub-directories below gallery/)
            # Gallery.report_galeries()
            # Gallery.report_gallery_items(GalleryKind.SHAPES)
            # return
            # list all the items for the Diagrams theme
            try:
                # itm = Gallery.find_gallery_item(
                #     GalleryKind.SHAPES,
                #     "Sun",
                #     search_match=SearchMatchKind.FULL,
                #     search_kind=SearchByKind.TITLE,
                # )
                # Gallery.report_gallery_item(itm)
                # item_type = Gallery.get_item_type(itm)
                # print(item_type)
                gs = ("Sun", "Callout-6", "8-Pointed-Star", "Shape-3")
                for g in gs:
                    graphic = Gallery.find_gallery_graphic(g)

                    Info.show_interfaces(f"Graphic: {g}", graphic)
            except GalleryNotFoundError:
                print("Gallery Item not found")


if __name__ == "__main__":
    GalleryInfo.main()
