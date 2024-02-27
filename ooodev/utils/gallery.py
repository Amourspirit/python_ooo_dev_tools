# PM - Notes, Oct 25, 2022
#
# this module has issues.
# find_gallery_item() returns XGalleryItem which in may have properties such as Graphic (XGraphic).
# the issues is as soon as the XGalleryItem is returned from find_gallery_item() it has already lost its properties.
# see my post: https://ask.libreoffice.org/t/libreoffice-wiping-object-properties-issue-with-gallerythemeprovider/83182
#
# Checking the reference count inside of find_gallery_item() show there is only 1 ref could. Which means 0 references
# because sys.getrefcount() add 1 reference to the object that is being checked.
# find_gallery_graphic() has been added which duplicates the code of find_gallery_item() due to these issues.
# find_gallery_graphic() does successfully return XGraphic object, However I am not sure if this graphic can be used
# as it is. I tried inserting the XGraphic into a Draw XShape and putting it on the document, however it appears to
# always be the same graphic even though the criteria is change for find_gallery_graphic()
from __future__ import annotations
from typing import TYPE_CHECKING, Any, List, cast
from pathlib import Path
from typing import overload
import uno
from com.sun.star.gallery import XGalleryItem
from com.sun.star.gallery import XGalleryTheme
from com.sun.star.gallery import XGalleryThemeProvider
from com.sun.star.graphic import XGraphic


from ooodev.utils import file_io as mFileIo
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils import info as mInfo
from ooodev.exceptions import ex as mEx
from ooodev.utils.kind.search_match_kind import SearchMatchKind as SearchMatchKind
from ooodev.utils.kind.gallery_kind import GalleryKind as GalleryKind
from ooodev.utils.kind.gallery_search_by_kind import SearchByKind as SearchByKind

from ooodev.events.event_singleton import _Events
from ooodev.events.lo_named_event import LoNamedEvent

from ooo.dyn.gallery.gallery_item_type import GalleryItemTypeEnum as GalleryItemTypeEnum
from ooo.lo.gallery.gallery_item_type import GalleryItemType as GalleryItemType

from ooodev.meta.static_meta import StaticProperty, classproperty

if TYPE_CHECKING:
    from com.sun.star.lang import EventObject
    from com.sun.star.beans import XPropertySetInfo


class GalleryObj:
    """
    Represents Most properties of ``XGalleryItem``.
    An instance of this class is returned from :py:meth:`~.gallery.Gallery.find_gallery_obj` in place of ``XGalleryItem`` instance.
    This is due to a `bug <https://bugs.documentfoundation.org/show_bug.cgi?id=151932>`_ in ``LO 7.4``.
    """

    # special case for this class. Matching most of GalleryItem properties so
    # using Camel case
    def __init__(self, itm: XGalleryItem) -> None:
        self._graphic = mProps.Props.get(itm, "Graphic", None)
        self._drawing = mProps.Props.get(itm, "Drawing", None)
        url = cast(str, mProps.Props.get(itm, "URL", ""))
        self._url = Gallery.get_absolute_url(url)
        self._gallery_item_type = mProps.Props.get(itm, "GalleryItemType", 0)
        self._title = mProps.Props.get(itm, "Title", "")
        self._implementation_id = getattr(itm, "ImplementationId", None)
        self._implementation_name = getattr(itm, "ImplementationName", "")
        self._property_set_info = getattr(itm, "PropertySetInfo", None)
        self._property_to_default = getattr(itm, "PropertyToDefault", None)
        self._thumbnail = mProps.Props.get(itm, "Thumbnail", None)

    @property
    def Graphic(self) -> XGraphic:
        return self._graphic

    @property
    def Drawing(self) -> Any:
        return self._drawing

    @property
    def URL(self) -> str:
        return self._url

    @property
    def GalleryItemType(self) -> int:
        return self._gallery_item_type

    @property
    def Title(self) -> str:
        return self._title

    @property
    def ImplementationId(self) -> uno.ByteSequence:
        return self._implementation_id  # type: ignore

    @property
    def ImplementationName(self) -> str:
        return self._implementation_name

    @property
    def PropertySetInfo(self) -> XPropertySetInfo:
        return self._property_set_info  # type: ignore

    @property
    def PropertyToDefault(self) -> Any:
        return self._property_to_default

    @property
    def Thumbnail(self) -> Any:
        return self._thumbnail


class Gallery(metaclass=StaticProperty):
    # region get_gallery()
    @overload
    @staticmethod
    def get_gallery(name: GalleryKind) -> XGalleryTheme: ...

    @overload
    @staticmethod
    def get_gallery(name: str) -> XGalleryTheme: ...

    @staticmethod
    def get_gallery(name: GalleryKind | str) -> XGalleryTheme:
        """
        Gets Gallery Theme

        Args:
            name (GalleryKind | str): Kind of gallery to get.

        Raises:
            GalleryError: If unable to get gallery theme.

        Returns:
            XGalleryTheme: Gallery Theme
        """
        # ensure name is GalleryKind | str,
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
            raise mEx.GalleryError(f'Error occurred getting gallery for "{name}"') from e

    # endregion get_gallery()

    @staticmethod
    def get_item_type_str(item_type: GalleryItemTypeEnum) -> str:
        """
        Gets string representation of ``GalleryItemTypeEnum``

        Args:
            item_type (GalleryItemTypeEnum): item

        Returns:
            str: string representation
        """
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
    def get_item_type(item: XGalleryItem | GalleryObj) -> GalleryItemTypeEnum:
        """
        Get item type

        Args:
            item (XGalleryItem | GalleryObj): item

        Raises:
            GalleryError: If error occurs.
        Returns:
            GalleryItemTypeEnum: Item type as enum.
        """
        try:
            if isinstance(item, GalleryObj):
                return GalleryItemTypeEnum(int(item.GalleryItemType))
            item_int_type = int(mProps.Props.get(item, "GalleryItemType"))
            return GalleryItemTypeEnum(item_int_type)
        except Exception as e:
            raise mEx.GalleryError("Error getting item type from gallery item") from e

    @classmethod
    def get_gallery_graphic(cls, item: XGalleryItem | GalleryObj) -> XGraphic:
        """
        Gets graphic form item

        Args:
            item (XGalleryItem | GalleryObj): Gallery Item

        Raises:
            GalleryError: If any other error occurs.

        Returns:
            XGraphic: Graphic object
        """
        try:
            if item is None:
                raise TypeError("item is expected to be of Type XGalleryItem, got None")
            # itm_type = cls.get_item_type(item)
            # if itm_type != GalleryItemTypeEnum.GRAPHIC:
            #     raise mEx.GalleryTypeError(
            #         f" Expected item to be GalleryItemType.DRAWING but got GalleryItemType.{itm_type.name}"
            #     )
            if isinstance(item, GalleryObj):
                result = item.Graphic
            else:
                result = mProps.Props.get(item, "Graphic")
            return mLo.Lo.qi(XGraphic, result, True)
        except mEx.GalleryTypeError:
            raise
        except Exception as e:
            raise mEx.GalleryError("Error getting gallery graphic") from e

    @staticmethod
    def get_gallery_path(item: XGalleryItem | GalleryObj) -> Path:
        """
        Gets gallery path

        Args:
            item (XGalleryItem | GalleryObj): Gallery Item

        Raises:
            GalleryError: If error occurs

        Returns:
            Path: Gallery path.
        """
        if item is None:
            raise TypeError("item is expected to be of Type XGalleryItem, got None")
        try:
            if isinstance(item, GalleryObj):
                url = item.URL
            else:
                url = str(mProps.Props.get(item, "URL", ""))
            if not url:
                raise mEx.GalleryError("Error getting gallery path")
            return mFileIo.FileIO.url_to_path(url)
        except mEx.GalleryError:
            raise
        except Exception as e:
            raise mEx.GalleryError("Error getting gallery path") from e

    # region find_gallery_item()
    @overload
    @classmethod
    def find_gallery_obj(cls, gallery_name: GalleryKind, name: str) -> GalleryObj: ...

    @overload
    @classmethod
    def find_gallery_obj(cls, gallery_name: str, name: str) -> GalleryObj: ...

    @overload
    @classmethod
    def find_gallery_obj(cls, gallery_name: GalleryKind, name: str, search_match: SearchMatchKind) -> GalleryObj: ...

    @overload
    @classmethod
    def find_gallery_obj(cls, gallery_name: str, name: str, search_match: SearchMatchKind) -> GalleryObj: ...

    @overload
    @classmethod
    def find_gallery_obj(cls, gallery_name: GalleryKind, name: str, *, search_kind: SearchByKind) -> GalleryObj: ...

    @overload
    @classmethod
    def find_gallery_obj(cls, gallery_name: str, name: str, *, search_kind: SearchByKind) -> GalleryObj: ...

    @overload
    @classmethod
    def find_gallery_obj(
        cls, gallery_name: str, name: str, search_match: SearchMatchKind, search_kind: SearchByKind
    ) -> GalleryObj: ...

    @overload
    @classmethod
    def find_gallery_obj(
        cls, gallery_name: GalleryKind, name: str, search_match: SearchMatchKind, search_kind: SearchByKind
    ) -> GalleryObj: ...

    @classmethod
    def find_gallery_obj(
        cls,
        gallery_name: GalleryKind | str,
        name: str,
        search_match: SearchMatchKind = SearchMatchKind.PARTIAL_IGNORE_CASE,
        search_kind: SearchByKind = SearchByKind.TITLE,
    ) -> GalleryObj:
        """
        Finds a Gallery Item

        Args:
            gallery_name (GalleryKind | str): Kind of gallery to search in.
            name (str): Name of item to look for. Could be a partial or full name of a title or a path. Determined by ``search_kind``.
            search_match (SearchMatchKind, optional): Search match option. Defaults to ``SearchMatchKind.PARTIAL_IGNORE_CASE``.
            search_kind (SearchByKind, optional): Determines what part of Gallery Item to search. Defaults to ``SearchByKind.TITLE``.

        Raises:
            GalleryNotFoundError: If Gallery Item is not found.
            GalleryError: If any other error occurs.

        Returns:
            GalleryObj: Gallery Item
        """
        # sourcery skip: low-code-quality, merge-duplicate-blocks
        result = None
        try:
            if search_match in (SearchMatchKind.FULL_IGNORE_CASE, SearchMatchKind.PARTIAL_IGNORE_CASE):
                nm = name.casefold()
                case_sensitive = False
            else:
                nm = name
                case_sensitive = True

            partial = search_match in (SearchMatchKind.PARTIAL, SearchMatchKind.PARTIAL_IGNORE_CASE)
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
                    url = str(mProps.Props.get(item, "URL", ""))
                    if not url:
                        continue
                    url = cls.get_absolute_url(url)
                    if url.startswith("private:"):
                        continue
                    fnm = mFileIo.FileIO.get_fnm((mFileIo.FileIO.url_to_path(url)))
                    match_str = fnm if case_sensitive else fnm.casefold()
                    if partial and nm in match_str:
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
                    if partial and nm in match_str:
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
                f'Error occurred trying to find in gallery: "{gallery_name}" for item: "{name}"'
            ) from e
        if result is None:
            raise mEx.GalleryNotFoundError(f'Not found. Tried to find in gallery: "{gallery_name}" for item: "{name}"')

        # cls.report_gallery_item(result)
        return GalleryObj(result)

    # endregion find_gallery_item()

    # region report_gallery_item()

    @classmethod
    def report_gallery_item(cls, item: XGalleryItem | GalleryObj) -> None:
        """
        Displays gallery information in the console.

        Args:
            item (XGalleryItem | GalleryObj): Gallery Item
        """
        if isinstance(item, GalleryObj):
            return cls._report_gallery_obj(item)
        return cls._report_gallery_item(item)

    @classmethod
    def _report_gallery_item(cls, item: XGalleryItem) -> None:
        if item is None:
            print("Gallery item is null")
            return

        print("Gallery item information:")
        url = None
        try:
            url = mProps.Props.get(item, "URL")
            if url is None:
                print("  URL: Value is None")
                print("  Fnm: Unable to compute due to no URL is None")
                print("  Path: Unable to compute due do no URL is None")
            else:
                url = Gallery.get_absolute_url(url)
                path = mFileIo.FileIO.uri_to_path(uri_fnm=url, ensure_absolute=False)
                print(f'  URL: "{url}"')
                print(f'  Fnm: "{mFileIo.FileIO.get_fnm(path)}"')
                print(f"  Path: {path}")
        except mEx.PropertyNotFoundError:
            print("  URL: Property NOT Found")
            print("  Fnm: Unable to compute due to no URL Property")
            print("  Path: Unable to compute due do no URL Property")
        except mEx.ConvertPathError:
            if url is not None:
                print(f'  URL: "{url}"')
            print("  Fnm: Unable to compute due to URL conversion error")
            print("  Path: Unable to compute due do URL conversion error")

        print(f'  Title: "{mProps.Props.get(item, "Title", "TITLE NOT FOUND")}"')
        try:
            item_int_type = int(mProps.Props.get(item, "GalleryItemType", GalleryItemTypeEnum.EMPTY.value))
            print(f"  Type: {cls.get_item_type_str(GalleryItemTypeEnum(item_int_type))}")
        except mEx.PropertyNotFoundError:
            print("  Type: Property NOT Found")

    @classmethod
    def _report_gallery_obj(cls, item: GalleryObj) -> None:
        if item is None:
            print("Gallery item is null")
            return

        print("Gallery item information:")
        url = None
        try:
            url = item.URL
            if url is None:
                print("  URL: Value is None")
                print("  Fnm: Unable to compute due to no URL is None")
                print("  Path: Unable to compute due do no URL is None")
            else:
                path = mFileIo.FileIO.uri_to_path(uri_fnm=url, ensure_absolute=False)
                print(f'  URL: "{url}"')
                print(f'  Fnm: "{mFileIo.FileIO.get_fnm(path)}"')
                print(f"  Path: {path}")
        except mEx.PropertyNotFoundError:
            print("  URL: Property NOT Found")
            print("  Fnm: Unable to compute due to no URL Property")
            print("  Path: Unable to compute due do no URL Property")
        except mEx.ConvertPathError:
            if url is not None:
                print(f'  URL: "{url}"')
            print("  Fnm: Unable to compute due to URL conversion error")
            print("  Path: Unable to compute due do URL conversion error")

        print(f'  Title: "{item.Title}"')
        try:
            item_int_type = item.GalleryItemType
            print(f"  Type: {cls.get_item_type_str(GalleryItemTypeEnum(item_int_type))}")
        except mEx.PropertyNotFoundError:
            print("  Type: Property NOT Found")

    # endregion report_gallery_item()

    @staticmethod
    def get_gallery_names() -> List[str]:
        """
        Gets a list of Gallery Names

        Returns:
            List[str]: Gallery names list
        """
        gtp = mLo.Lo.create_instance_mcf(
            XGalleryThemeProvider, "com.sun.star.gallery.GalleryThemeProvider", raise_err=True
        )
        themes = gtp.getElementNames()
        return sorted(themes)

    @staticmethod
    def report_galleries() -> None:
        """
        Displays a list of galleries in the console
        """
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
        """
        Displays Gallery Items information in console for a Gallery.

        Args:
            gallery_name (GalleryKind | str): Gallery kind
        """
        try:
            gallery = cls.get_gallery(gallery_name)
            num_pics = gallery.getCount()
            print(f'Gallery: "{gallery.getName()}" ({num_pics})')
            for i in range(num_pics):
                try:
                    item = GalleryObj(mLo.Lo.qi(XGalleryItem, gallery.getByIndex(i), True))
                    cls._report_gallery_obj(item)
                    # url = str(mProps.Props.get(item, "URL"))
                    # print(f"  {mFileIo.FileIO.get_fnm((mFileIo.FileIO.url_to_path(url)))}")
                except Exception as ex:
                    print(f"Could not access gallery item: {i}")
                    print(f"  {ex}")
        except Exception as e:
            print("Error reporting gallery items")
            print(f"  {e}")
        print()

    @staticmethod
    def get_absolute_url(url: str) -> str:
        """
        Get absolute url
        Some url's may have relative paths in them.

        Args:
            url (str): url

        Returns:
            str: absolute url

        Note:
            If a ``url`` is not a file path such as ``private:gallery/svdraw/dd2051``
            then it is returned verbatim.
        """
        # https://tinyurl.com/2q3evk44
        if url:
            if url.startswith("private:"):
                result = url
            else:
                try:
                    # convert url to path to ensure is a valid format.
                    _ = mFileIo.FileIO.uri_to_path(uri_fnm=url, ensure_absolute=False)
                    result = mFileIo.FileIO.uri_absolute(url)
                except mEx.ConvertPathError:
                    # could be format such as "private:gallery/svdraw/dd2051"
                    result = url
        else:
            result = ""
        return result

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
        data_attrs = ("_gallery_dir",)
        for attr in data_attrs:
            if hasattr(Gallery, attr):
                delattr(Gallery, attr)


_Events().on(LoNamedEvent.BRIDGE_DISPOSED, _GalleryManager.on_disposed)


__all__ = ("Gallery",)
