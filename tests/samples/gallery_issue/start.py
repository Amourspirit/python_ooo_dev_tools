#!/usr/bin/env python3
from __future__ import annotations
import sys
import uno
from com.sun.star.gallery import XGalleryItem
from ooodev.loader.lo import Lo
from gallery import Gallery


def find_gallery_item(gallery_name: str, name: str) -> XGalleryItem:
    gallery = Gallery.get_gallery(gallery_name)  # XGalleryTheme
    num_pics = gallery.getCount()
    result = None
    for i in range(num_pics):
        item = gallery.getByIndex(i)
        # run code to find item.
        result = item
        break

    # sys.getrefcount() adds a refcount so sys.getrefcount() - 1
    # Reference count for result.Drawing is 0
    print()
    print("result.Drawing ref count:", sys.getrefcount(result.Drawing) - 1)
    print()
    # this means as soon as result is return from this function then
    # python garbage collector deletes result.Drawing

    # this next line report as expected. result.Drawing is XDrawing instance
    Gallery.report_gallery_item(result)
    return result


def main() -> int:
    with Lo.Loader(Lo.ConnectSocket(), opt=Lo.Options(verbose=True)):
        item = find_gallery_item(gallery_name="Shapes", name="Sun")
        print()
        print("item.Drawing type reported on function result", type(item.Drawing))
        print()
        # Reference count for item.Drawing is 0
        Gallery.report_gallery_item(item)  # at this point item.Drawing is None
        return 0


if __name__ == "__main__":
    SystemExit(main())
