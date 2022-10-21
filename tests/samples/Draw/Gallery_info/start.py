from __future__ import annotations
import uno
from ooodev.utils.lo import Lo
from ooodev.utils.gallery import Gallery, SearchMatchKind, GalleryKind, SearchByKind
from ooodev.exceptions.ex import GalleryNotFoundError


class GalleryInfo:
    @staticmethod
    def main() -> None:
        with Lo.Loader(Lo.ConnectPipe()) as loader:

            # list all the gallery themes (i.e. the sub-directories below gallery/)
            # Gallery.report_galeries()

            # list all the items for the Diagrams theme
            try:
                itm = Gallery.find_gallery_item(
                    GalleryKind.DIAGRAMS,
                    "Pyramid-Fancy",
                    search_match=SearchMatchKind.FULL,
                    search_kind=SearchByKind.TITLE,
                )
                Gallery.report_gallery_item(itm)
                # item_type = Gallery.get_item_type(itm)
                # print(item_type)
                # graphic = Gallery.get_gallery_graphic(item=itm)
                # print(graphic)
            except GalleryNotFoundError:
                print("Gallery Item not found")


if __name__ == "__main__":
    GalleryInfo.main()
