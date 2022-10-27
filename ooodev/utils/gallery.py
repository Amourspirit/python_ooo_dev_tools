# PM - Notes, Oct 25, 2022
#
# this module has issues.
# find_gallery_item() returns XGalleryItem which in may have properties such as Graphic (XGraphic).
# the issues is as soon as the XGalleryItem is returned from find_gallery_item() it has already lost its properties.
# see my post: https://ask.libreoffice.org/t/libreoffice-wiping-object-properties-issue-with-gallerythemeprovider/83182
#
# Checking the reference count inside of find_gallery_item() show there is only 1 ref could. Which means 0 references
# becuase sys.getrefcount() add 1 reference to the object that is being checked.
# find_gallery_graphic() has been added which duplicates the code of find_gallery_item() due to these issues.
# find_gallery_graphic() does successfully return XGraphic object, However I am not sure if this graphic can be used
# as it is. I tried inserting the XGraphic into a Draw XShape and putting it on the document, however it appears to
# always be the same graphic even though the criteria is change for find_gallery_graphic()

from __future__ import annotations
from typing import TYPE_CHECKING, Any
from pathlib import Path
from typing import overload
import uno
from com.sun.star.gallery import XGalleryItem
from com.sun.star.gallery import XGalleryTheme
from com.sun.star.gallery import XGalleryThemeProvider
from com.sun.star.graphic import XGraphic

from . import file_io as mFileIo
from . import lo as mLo
from . import props as mProps
from . import info as mInfo
from ..exceptions import ex as mEx
from .kind.search_match_kind import SearchMatchKind as SearchMatchKind
from .kind.gallery_kind import GalleryKind as GalleryKind
from .kind.gallery_search_by_kind import SearchByKind as SearchByKind

from ..events.event_singleton import _Events
from ..events.lo_named_event import LoNamedEvent

from ooo.dyn.gallery.gallery_item_type import GalleryItemTypeEnum as GalleryItemTypeEnum
from ooo.lo.gallery.gallery_item_type import GalleryItemType as GalleryItemType

from ..meta.static_meta import StaticProperty, classproperty

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject


class Gallery(metaclass=StaticProperty):
    # region get_gallery()
    @overload
    @staticmethod
    def get_gallery(name: GalleryKind) -> XGalleryTheme:
        ...

    @overload
    @staticmethod
    def get_gallery(name: str) -> XGalleryTheme:
        ...

    @staticmethod
    def get_gallery(name: GalleryKind | str) -> XGalleryTheme:

        # enusre name is GalleryKind | str,
        mInfo.Info.is_type_enum_multi(alt_type="str", enum_type=GalleryKind, enum_val=name, arg_name="name")

        try:
            gtp = mLo.Lo.create_instance_mcf(
                XGalleryThemeProvider, "com.sun.star.gallery.GalleryThemeProvider", raise_err=True
            )
            obj = gtp.getByName(str(name))
            if obj is None:
                raise mEx.UnKnownError("None Value: getByName() returned None value")
            return mLo.Lo.qi(XGalleryTheme, obj, True)
        except Exception as e:
            raise mEx.GalleryError(f'Error occured getting gallery for "{name}"') from e

    # endregion get_gallery()

    @staticmethod
    def get_item_type_str(item_type: GalleryItemTypeEnum) -> str:
        if item_type == GalleryItemTypeEnum.EMPTY:
            return "empty"
        elif item_type == GalleryItemTypeEnum.GRAPHIC:
            return "graphic"
        elif item_type == GalleryItemTypeEnum.MEDIA:
            return "media"
        elif item_type == GalleryItemType.DRAWING:
            return "drawing"
        return "??"  # just in case enum change in future.

    @staticmethod
    def get_item_type(item: XGalleryItem) -> GalleryItemTypeEnum:
        try:
            item_int_type = int(mProps.Props.get(item, "GalleryItemType"))
            return GalleryItemTypeEnum(item_int_type)
        except Exception as e:
            raise mEx.GalleryError("Error getting item type from gallery item") from e

    @classmethod
    def get_gallery_graphic(cls, item: XGalleryItem) -> XGraphic:
        """
        Gets graphic form item

        Args:
            item (XGalleryItem): Gallery Item

        Raises:
            GalleryTypeError: If ``item.Graphic`` is not ``GalleryItemType.DRAWING``
            GalleryError: If any other error occurs.

        Returns:
            XGraphic: Graphic object
        """
        try:
            if item is None:
                raise TypeError("item is expected to be of Type XGalleryItem, got None")
            itm_type = cls.get_item_type(item)
            if itm_type != GalleryItemTypeEnum.DRAWING:
                raise mEx.GalleryTypeError(
                    f" Expected item to be GalleryItemType.DRAWING but got GalleryItemType.{itm_type.name}"
                )
            result = mProps.Props.get(item, "Graphic")
            return mLo.Lo.qi(XGraphic, result, True)
        except mEx.GalleryTypeError:
            raise
        except Exception as e:
            raise mEx.GalleryError("Error getting gallery graphic") from e

    @staticmethod
    def get_gallery_path(item: XGalleryItem) -> Path:
        if item is None:
            raise TypeError("item is expected to be of Type XGalleryItem, got None")
        try:
            url = str(mProps.Props.get(item, "URL"))
            return mFileIo.FileIO.url_to_path(url)
        except Exception as e:
            raise mEx.GalleryError("Error getting gallery path") from e

    # region find_gallery_item()
    @overload
    @classmethod
    def find_gallery_item(cls, gallery_name: str, name: str) -> XGalleryItem:
        ...

    @overload
    @classmethod
    def find_gallery_item(cls, gallery_name: GalleryKind, name: str) -> XGalleryItem:
        ...

    @overload
    @classmethod
    def find_gallery_item(cls, gallery_name: str, name: str, search_match: SearchMatchKind) -> XGalleryItem:
        ...

    @overload
    @classmethod
    def find_gallery_item(cls, gallery_name: GalleryKind, name: str, search_match: SearchMatchKind) -> XGalleryItem:
        ...

    @overload
    @classmethod
    def find_gallery_item(cls, gallery_name: str, name: str, *, search_kind: SearchByKind) -> XGalleryItem:
        ...

    @overload
    @classmethod
    def find_gallery_item(cls, gallery_name: GalleryKind, name: str, *, search_kind: SearchByKind) -> XGalleryItem:
        ...

    @overload
    @classmethod
    def find_gallery_item(
        cls, gallery_name: str, name: str, search_match: SearchMatchKind, search_kind: SearchByKind
    ) -> XGalleryItem:
        ...

    @overload
    @classmethod
    def find_gallery_item(
        cls, gallery_name: GalleryKind, name: str, search_match: SearchMatchKind, search_kind: SearchByKind
    ) -> XGalleryItem:
        ...

    @classmethod
    def find_gallery_item(
        cls,
        gallery_name: GalleryKind | str,
        name: str,
        search_match: SearchMatchKind = SearchMatchKind.PARTIAL_IGNORE_CASE,
        search_kind: SearchByKind = SearchByKind.TITLE,
    ) -> XGalleryItem:

        result = None
        try:
            if search_match == SearchMatchKind.FULL_IGNORE_CASE or search_kind == SearchMatchKind.PARTIAL_IGNORE_CASE:
                nm = name.casefold()
                case_sensitive = False
            else:
                nm = name
                case_sensitive = True

            if search_match == SearchMatchKind.PARTIAL or search_match == SearchMatchKind.PARTIAL_IGNORE_CASE:
                partial = True
            else:
                partial = False

            gallery = cls.get_gallery(gallery_name)
            num_pics = gallery.getCount()
            mLo.Lo.print(f'Searching gallery "{gallery.getName()}" for "{name}"')
            if case_sensitive:
                mLo.Lo.print("  Search is case sensitive")
            else:
                mLo.Lo.print("  Search is ignoring case")
            if partial:
                mLo.Lo.print("  Searching for a partial match")
            else:
                mLo.Lo.print("  Searching for a full match")
            mLo.Lo.print()

            for i in range(num_pics):
                item = mLo.Lo.qi(XGalleryItem, gallery.getByIndex(i), True)
                if search_kind == SearchByKind.FILE_NAME:
                    url = str(mProps.Props.get(item, "URL"))
                    fnm = mFileIo.FileIO.get_fnm((mFileIo.FileIO.url_to_path(url)))
                    match_str = fnm if case_sensitive else fnm.casefold()
                    if partial and match_str in nm:
                        mLo.Lo.print(f"Found matching item: {fnm}")
                        result = item
                        break
                    else:
                        if match_str == nm:
                            mLo.Lo.print(f"Found matching item: {fnm}")
                            result = item
                            break
                elif search_kind == SearchByKind.TITLE:
                    title = str(mProps.Props.get(item, "Title"))
                    match_str = title if case_sensitive else title.casefold()
                    if partial and match_str in nm:
                        mLo.Lo.print(f"Found matching item: {title}")
                        result = item
                        break
                    else:
                        if match_str == nm:
                            mLo.Lo.print(f"Found matching item: {title}")
                            result = item
                            break

        except Exception as e:
            raise mEx.GalleryError(
                f'Error occured trying to find in gallery: "{gallery_name}" for item: "{name}"'
            ) from e
        if result is None:
            raise mEx.GalleryNotFoundError(f'Not found. Tried to find in gallery: "{gallery_name}" for item: "{name}"')
        cls.report_gallery_item(result)
        return result

    # endregion find_gallery_item()

    # region find_gallery_graphic()
    @overload
    @classmethod
    def find_gallery_graphic(cls, name: str) -> XGraphic:
        ...

    @overload
    @classmethod
    def find_gallery_graphic(cls, name: str, search_match: SearchMatchKind) -> XGraphic:
        ...

    @overload
    @classmethod
    def find_gallery_graphic(cls, name: str, *, search_kind: SearchByKind) -> XGraphic:
        ...

    @overload
    @classmethod
    def find_gallery_graphic(cls, name: str, search_match: SearchMatchKind, search_kind: SearchByKind) -> XGraphic:
        ...

    @classmethod
    def find_gallery_graphic(
        cls,
        name: str,
        search_match: SearchMatchKind = SearchMatchKind.PARTIAL_IGNORE_CASE,
        search_kind: SearchByKind = SearchByKind.TITLE,
    ) -> XGraphic:

        result = None
        try:
            if search_match == SearchMatchKind.FULL_IGNORE_CASE or search_kind == SearchMatchKind.PARTIAL_IGNORE_CASE:
                nm = name.casefold()
                case_sensitive = False
            else:
                nm = name
                case_sensitive = True

            if search_match == SearchMatchKind.PARTIAL or search_match == SearchMatchKind.PARTIAL_IGNORE_CASE:
                partial = True
            else:
                partial = False

            gallery = cls.get_gallery(GalleryKind.SHAPES)
            num_pics = gallery.getCount()
            mLo.Lo.print(f'Searching gallery "{gallery.getName()}" for "{name}"')
            if case_sensitive:
                mLo.Lo.print("  Search is case sensitive")
            else:
                mLo.Lo.print("  Search is ignoring case")
            if partial:
                mLo.Lo.print("  Searching for a partial match")
            else:
                mLo.Lo.print("  Searching for a full match")
            mLo.Lo.print()

            for i in range(num_pics):
                item = mLo.Lo.qi(XGalleryItem, gallery.getByIndex(i), True)
                if search_kind == SearchByKind.FILE_NAME:
                    url = str(mProps.Props.get(item, "URL"))
                    fnm = mFileIo.FileIO.get_fnm((mFileIo.FileIO.url_to_path(url)))
                    match_str = fnm if case_sensitive else fnm.casefold()
                    if partial and match_str in nm:
                        mLo.Lo.print(f"Found matching item: {fnm}")
                        result = item
                        break
                    else:
                        if match_str == nm:
                            mLo.Lo.print(f"Found matching item: {fnm}")
                            result = item
                            break
                elif search_kind == SearchByKind.TITLE:
                    title = str(mProps.Props.get(item, "Title"))
                    match_str = title if case_sensitive else title.casefold()
                    if partial and match_str in nm:
                        mLo.Lo.print(f"Found matching item: {title}")
                        result = item
                        break
                    else:
                        if match_str == nm:
                            mLo.Lo.print(f"Found matching item: {title}")
                            result = item
                            break

        except Exception as e:
            raise mEx.GalleryError(
                f'Error occured trying to find in gallery: "{GalleryKind.SHAPES}" for item: "{name}"'
            ) from e
        # Gallery._TMP_FIND = result.Graphic
        if result is None:
            raise mEx.GalleryNotFoundError(
                f'Not found. Tried to find in gallery: "{GalleryKind.SHAPES}" for item: "{name}"'
            )
        try:
            graphic = mProps.Props.get(result, "Graphic")
        except mEx.PropertyNotFoundError as e:
            raise mEx.GalleryNotFoundError("Error getting Graphic Property") from e
        return graphic

    # endregion find_gallery_graphic()

    @classmethod
    def report_gallery_item(cls, item: XGalleryItem) -> None:
        if item is None:
            print("Gallery item is null")
            return

        print("Gallery item information:")
        try:
            url = str(mProps.Props.get(item, "URL"))
            path = mFileIo.FileIO.uri_to_path(uri_fnm=url, ensure_absolute=False)
            print(f'  URL: "{url}"')
            print(f'  Fnm: "{Path(Gallery.gallery_dir, mFileIo.FileIO.get_fnm(path))}"')
            print(f"  Path: {path}")
        except mEx.PropertyNotFoundError:
            print("  URL: Property NOT Found")
            print("  Fnm: Unable to compute due to no URL Property")
            print("  Path: Unable to compute due do no URL Property")
        try:
            print(f'  Title: "{mProps.Props.get(item, "Title")}"')
        except mEx.PropertyNotFoundError:
            print("  Title: Property NOT Found")
        try:
            item_int_type = int(mProps.Props.get(item, "GalleryItemType"))
            print(f"  Type: {cls.get_item_type_str(GalleryItemTypeEnum(item_int_type))}")
        except mEx.PropertyNotFoundError:
            print("  Type: Property NOT Found")

    @staticmethod
    def report_galeries() -> None:
        gtp = mLo.Lo.create_instance_mcf(
            XGalleryThemeProvider, "com.sun.star.gallery.GalleryThemeProvider", raise_err=True
        )
        themes = gtp.getElementNames()
        print(f"No. of themes: {len(themes)}")
        for theme in themes:
            try:
                gallery = mLo.Lo.qi(XGalleryTheme, gtp.getByName(theme), True)
                print(f'  "{gallery.getName()}" ({gallery.getCount()})')
            except Exception as e:
                print(f'Could not access gallery for "{theme}"')
                print(f"  {e}")
        print()

    @classmethod
    def report_gallery_items(cls, gallery_name: GalleryKind | str) -> None:
        try:
            gallery = cls.get_gallery(gallery_name)
            num_pics = gallery.getCount()
            print(f'Gallery: "{gallery.getName()}" ({num_pics})')
            for i in range(num_pics):
                try:
                    item = mLo.Lo.qi(XGalleryItem, gallery.getByIndex(i), True)
                    cls.report_gallery_item(item)
                    # url = str(mProps.Props.get(item, "URL"))
                    # print(f"  {mFileIo.FileIO.get_fnm((mFileIo.FileIO.url_to_path(url)))}")
                except Exception as ex:
                    print(f"Could not access galery item: {i}")
                    print(f"  {ex}")
        except Exception as e:
            print("Error reporting gallery items")
            print(f"  {e}")
        print()

    @classproperty
    def gallery_dir(cls) -> Path:
        """
        Get the first directory that contain the Gallery database and multimedia files.
        """
        try:
            return cls._gallery_dir
        except AttributeError:
            cls._gallery_dir = mInfo.Info.get_gallery_dir()
        return cls._gallery_dir


class _GalleryManager:
    """Manages clearing and resetting for Gallery static class"""

    @staticmethod
    def on_disposed(source: Any, event: EventObject) -> None:
        # Clean up static properties that may have been dynamically created.
        # print("Gallery Static Property Cleanup")
        dattrs = ("_gallery_dir",)
        for attr in dattrs:
            if hasattr(Gallery, attr):
                delattr(Gallery, attr)


_Events().on(LoNamedEvent.BRIDGE_DISPOSED, _GalleryManager.on_disposed)


__all__ = ("Gallery",)
