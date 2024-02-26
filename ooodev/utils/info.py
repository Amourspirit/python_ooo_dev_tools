# coding: utf-8
# Python conversion of Info.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
from __future__ import annotations
import sys
import contextlib
from enum import Enum, IntFlag
from pathlib import Path
import mimetypes
from typing import TYPE_CHECKING, Any, Tuple, List, cast, overload, Optional, Set
from typing import Type, TypeVar, Union

import uno

from com.sun.star.awt import XToolkit
from com.sun.star.beans import XHierarchicalPropertySet
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XContentEnumerationAccess
from com.sun.star.container import XNameAccess
from com.sun.star.container import XNameContainer
from com.sun.star.deployment import XPackageInformationProvider
from com.sun.star.document import XDocumentPropertiesSupplier
from com.sun.star.document import XTypeDetection
from com.sun.star.frame import XModule
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XTypeProvider
from com.sun.star.reflection import XIdlReflection
from com.sun.star.style import XStyleFamiliesSupplier
from com.sun.star.util import XChangesBatch


from ooo.dyn.beans.property_value import PropertyValue
from ooo.dyn.beans.property_concept import PropertyConceptEnum
from ooo.dyn.beans.the_introspection import theIntrospection
from ooo.dyn.lang.locale import Locale  # struct

from ooodev.utils.decorator.deprecated import deprecated
from ooodev.loader.inst.service import Service as LoService
from ooodev.utils import date_time_util as mDate
from ooodev.utils import file_io as mFileIO
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.units import unit_convert as mConvert
from ooodev.events.args.event_args import EventArgs
from ooodev.events.event_singleton import _Events
from ooodev.events.lo_named_event import LoNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.meta.static_meta import StaticProperty, classproperty
from ooodev.utils.kind.info_paths_kind import InfoPathsKind as InfoPathsKind
from ooodev.utils.sys_info import SysInfo
from ooodev.proto import uno_proto

if TYPE_CHECKING:
    from com.sun.star.awt import FontDescriptor
    from com.sun.star.beans import XPropertyContainer
    from com.sun.star.document import XDocumentProperties
    from com.sun.star.reflection import XIdlMethod
    from .type_var import PathOrStr
else:
    PathOrStr = Any

if sys.version_info >= (3, 10):
    from typing import TypeGuard, TypeAlias
else:
    from typing_extensions import TypeGuard, TypeAlias


_T = TypeVar("_T")
ClassInfo: TypeAlias = Union[Type[_T], Tuple["ClassInfo[_T]", ...]]


class Info(metaclass=StaticProperty):
    REG_MOD_FNM = "registrymodifications.xcu"
    NODE_PRODUCT = "/org.openoffice.Setup/Product"
    NODE_L10N = "/org.openoffice.Setup/L10N"
    NODE_OFFICE = "/org.openoffice.Setup/Office"

    NODE_PATHS = (NODE_PRODUCT, NODE_L10N)
    # MIME_FNM = "mime.types"

    # from https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Properties_of_a_Filter
    class Filter(IntFlag):
        """
        Every filter inside LibreOffice is specified by the properties of this enum.
        """

        IMPORT = 0x00000001
        """This filter is used only for internal purposes and so can be used only in API calls. Users won't see it ever. """
        EXPORT = 0x00000002
        """The filter supports the service com.sun.star.document.ExportFilter. It will be shown in the dialog "File-Export". If the filter also has the "IMPORT" flag set, it will be shown in the dialog "File-Save". This makes sure that a format that a user chooses in the save dialog can be loaded again. The same is not guaranteed for a format chosen in "File-Export". """
        TEMPLATE = 0x00000004
        """Filter denotes a template filter (means, by default all documents opened by it become an "untitled" one) """
        INTERNAL = 0x00000008
        """This filter is used only for internal purposes and so can be used only in API calls. Users won't see it ever. """
        TEMPLATEPATH = 0x00000010
        """Must always be set together with "TEMPLATE" to make this feature flag work; soon becoming deprecated"""
        OWN = 0x00000020
        """The filter is a native Apache OpenOffice format (ODF or otherwise)."""
        ALIEN = 0x00000040
        """The filter may lose some information upon saving. """
        DEFAULT = 0x00000100
        """This is the "best" filter for the document type it works on that is guaranteed not so lose any data on export. By default this filter will be used or suggested for every storing process unless the user has chosen a different default filter in the options dialog."""
        SUPPORTSSELECTION = 0x00000400
        """Filter can export only the selected part of a document. This information enables Apache OpenOffice to enable a corresponding check box in the "File-Export" dialog."""
        NOTINFILEDIALOG = 0x00001000
        """This filter will not be shown in the file dialog's filter list"""
        NOTINCHOOSER = 0x00002000
        """This filter will not be shown in the dialog box for choosing a filter in case Apache OpenOffice was not able to detect one"""
        READONLY = 0x00010000
        """All documents imported by this filter will automatically be in read-only state"""
        PREFERRED = 0x10000000
        """The filter is preferred in case of multiple filters for the same file type exist in the configuration"""
        THIRDPARTYFILTER = 0x00080000
        """
        The filter is a UNO component filter, as opposed to the internal C++ filters.
        This is an artifact that will vanish over time.

        AKA: 3RDPARTYFILTER
        """

    class RegPropKind(Enum):
        PROPERTY = 1
        VALUE = 2

    @staticmethod
    def get_fonts() -> Tuple[FontDescriptor, ...]:
        """
        Gets fonts.

        |lo_unsafe|

        Returns:
            Tuple[FontDescriptor, ...]: Font Descriptors
        """
        xtoolkit = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit", raise_err=True)
        device = xtoolkit.createScreenCompatibleDevice(0, 0)
        if device is None:
            mLo.Lo.print("Could not access graphical output device")
            return ()
        return device.getFontDescriptors()

    @classmethod
    def get_font_names(cls) -> List[str]:
        """
        Gets font names.

        |lo_unsafe|

        Returns:
            List[str]: Font names
        """
        fds = cls.get_fonts()
        if fds is None:
            return []

        names_set = {fd.Name for fd in fds}
        return sorted(names_set)

    @classmethod
    def get_font_descriptor(cls, name: str, style: str) -> FontDescriptor | None:
        """
        Gets font descriptor for a font name with a font style such as ``Bold Italic``.

        This method is useful for obtaining a font descriptor for a specific font name and style,
        which can be used for various purposes such as setting the font of text in a document or application.

        |lo_unsafe|

        Args:
            name (str): Font Name
            style (str): Font Style

        Returns:
            FontDescriptor: Instance if found; Otherwise None.

        .. versionadded:: 0.9.0
        """
        if not name:
            return None
        if not style:
            return None
        style = style.casefold()
        fds = cls.get_fonts()
        return next(
            (fd for fd in fds if fd.Name == name and fd.StyleName.casefold() == style),
            None,
        )

    @staticmethod
    def get_font_mono_name() -> str:
        """
        Gets a general font such as ``Courier New`` (windows) or ``Liberation Mono``.

        |lo_safe|

        Returns:
            str: Font Name

        See Also:
            `Fonts <https://wiki.documentfoundation.org/Fonts>`_ on Document Foundation’s wiki
        """
        pf = SysInfo.get_platform()
        if pf == SysInfo.PlatformEnum.WINDOWS:
            return "Courier New"
        else:
            return "Liberation Mono"  # Metrically compatible with Courier New

    @staticmethod
    def get_font_general_name() -> str:
        """
        Gets a general font such as ``Times New Roman`` (windows) or ``Liberation Serif``.

        |lo_safe|

        Returns:
            str: Font Name

        See Also:
            `Fonts <https://wiki.documentfoundation.org/Fonts>`_ on Document Foundation’s wiki
        """
        pf = SysInfo.get_platform()
        if pf == SysInfo.PlatformEnum.WINDOWS:
            return "Times New Roman"
        else:
            return "Liberation Serif"  # Metrically compatible with Times New Roman

    @classmethod
    def get_reg_mods_path(cls) -> str:
        """
        Get registered modifications path.

        |lo_unsafe|

        Returns:
            str: registered modifications path
        """
        user_cfg_dir = mFileIO.FileIO.url_to_path(cls.get_paths("UserConfig"))
        parent_path = user_cfg_dir.parent
        return str(parent_path / cls.REG_MOD_FNM)

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str) -> str: ...

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str) -> str: ...

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str, node: str) -> str: ...

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, *, kind: Info.RegPropKind) -> str: ...

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, *, kind: Info.RegPropKind, idx: int) -> str: ...

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str, *, idx: int) -> str: ...

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str, node: str, kind: Info.RegPropKind) -> str: ...

    @classmethod
    def get_reg_item_prop(
        cls,
        item: str,
        prop: str = "",
        node: Optional[str] = None,
        kind: Info.RegPropKind = RegPropKind.PROPERTY,
        idx: int = 0,
    ) -> str:
        """
        Gets value from ``registrymodifications.xcu``.

        |lo_safe|

        Args:
            item (str): item name
            prop (str): property value
            node (str): node
            kind (Info.RegPropKind): property or value
            idx: (int): index of value to return

        Raises:
            ValueError: if unable to get value

        Returns:
            str: value from ``registrymodifications.xcu``. e.g. ``Writer/MailMergeWizard``, ``None``, ``MailAddress``

        Note:
            Often is it not necessary to get value from registry.
            Which is not available in macro mode.

            Often the configuration option can be gotten via :py:meth:`~.info.Info.get_config`.

        See Also:
            :py:meth:`~.info.Info.get_config`
        """
        # sourcery skip: merge-else-if-into-elif, raise-specific-error, remove-unnecessary-else, swap-if-else-branches

        # return value from "registrymodifications.xcu"
        # windows C:\Users\user\AppData\Roaming\LibreOffice\4\user\registrymodifications.xcu
        # e.g. val = Info.get_reg_item_prop("Calc/Calculate/Other", "DecimalPlaces")
        # e.g. val = Info.get_reg_item_prop("Logging/Settings", "LogLevel", "org.openoffice.logging.sdbc.DriverManager")
        # e.g. val = Info.get_reg_item_prop("Writer/Layout/Other/TabStop", kind=Info.RegPropKind.VALUE, idx=1)
        # This xpath doesn't deal with all cases in the XCU file, which sometimes
        # has many node levels between the item and the prop
        if mLo.Lo.is_macro_mode:
            raise mEx.NotSupportedMacroModeError("get_reg_item_prop() is not supported from a macro")
        try:
            import xml.etree.ElementTree as ET  # type: ignore[import]
        except ImportError as e:
            raise Exception("get_reg_item_prop() requires xml python package") from e

        try:
            fnm = cls.get_reg_mods_path()
            tree = ET.parse(fnm)
            xpath = ""
            if not prop:
                kind = Info.RegPropKind.VALUE
            if node is None:
                if kind == Info.RegPropKind.PROPERTY:
                    xpath = f'.//item[@oor:path="/org.openoffice.Office.{item}"]/prop[@oor:name="{prop}"]/value'
                elif kind == Info.RegPropKind.VALUE:
                    xpath = f'.//item[@oor:path="/org.openoffice.Office.{item}"]/value'
            else:
                if kind == Info.RegPropKind.PROPERTY:
                    xpath = f'.//item[@oor:path="/org.openoffice.Office.{item}"]/node[@oor:name="{node}"]/prop[@oor:name="{prop}"]/value'
                elif kind == Info.RegPropKind.VALUE:
                    xpath = f".//item[/prop[@oor:name='{item}']/node[@oor:name='{node}']/value"
            if not xpath:
                raise ValueError("Unable to create xpath. Recheck arguments.")
            value = tree.findall(xpath, namespaces={"oor": "http://openoffice.org/2001/registry"})
            if not value:
                raise Exception("Item Property not found")
            else:
                try:
                    first = value[idx]
                except IndexError:
                    raise
                if first.text is None:
                    raise Exception("Item Property is None (?)")
                value = first.text.strip()
                if value == "":
                    raise Exception("Item Property is white space (?)")
            return value
        except Exception as e:
            raise ValueError("Unable to get value from registrymodifications.xcu") from e

    @overload
    @classmethod
    def get_config(cls, node_str: str) -> Any:
        """
        Get config.

        |lo_unsafe|

        Args:
            node_str (str): node string

        Raises:
            ConfigError: if unable to get config

        Returns:
            object: config
        """
        ...

    @overload
    @classmethod
    def get_config(cls, node_str: str, node_path: str) -> Any:
        """
        Get config.

        |lo_unsafe|

        Args:
            node_str (str): node string
            node_path (str): node_path

        Raises:
            ConfigError: if unable to get config

        Returns:
            object: config
        """
        ...

    @classmethod
    def get_config(cls, node_str: str, node_path: Optional[str] = None) -> Any:
        """
        Get config.

        |lo_unsafe|

        Args:
            node_str (str): node string
            node_path (str): node_path

        Raises:
            ConfigError: if unable to get config

        Returns:
            object: config

        See Also:
            :ref:`ch03`

        Example:
            .. code-block:: python

                props = mLo.Lo.qi(
                    XPropertySet,
                    mInfo.Info.get_config(node_str="Other", node_path="/org.openoffice.Office.Writer/Layout/"),
                )
                ts_val = props.getPropertyValue("TabStop") # int val
        """
        # props = Lo.qi(XPropertySet, Info.get_config(node_str='Data', node_path="/org.openoffice.UserProfile/"))
        # Props.show_obj_props("User Data", props)
        try:
            if node_path is None:
                return cls._get_config2(node_str=node_str)
            return cls._get_config1(node_str=node_str, node_path=node_path)
        except Exception as e:
            msg = f"Unable to get configuration for '{node_str}'"
            if node_path is not None:
                msg += f" with path: '{node_path}'"
            raise mEx.ConfigError(msg) from e

    @classmethod
    def _get_config1(cls, node_str: str, node_path: str):
        """LO UN-Safe Method"""
        props = cls.get_config_props(node_path)
        return mProps.Props.get(props, node_str)

    @classmethod
    def _get_config2(cls, node_str: str) -> Any:
        """LO UN-Safe Method"""
        for node_path in cls.NODE_PATHS:
            with contextlib.suppress(mEx.PropertyNotFoundError):
                return cls._get_config1(node_str=node_str, node_path=node_path)
        raise mEx.ConfigError(f"{node_str} not found in common node paths")

    @staticmethod
    def get_config_props(node_path: str) -> XPropertySet:
        """
        Get config properties.

        |lo_unsafe|

        Args:
            node_path (str): nod path

        Raises:
            PropertyError: if unable to get get property set

        Returns:
            XPropertySet: Property set
        """
        try:
            con_prov = mLo.Lo.create_instance_mcf(
                XMultiServiceFactory,
                "com.sun.star.configuration.ConfigurationProvider",
                raise_err=True,
            )
            p = mProps.Props.make_props(nodepath=node_path)
            ca = con_prov.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", p)
            return mLo.Lo.qi(XPropertySet, ca, True)
        except Exception as e:
            raise mEx.PropertyError(node_path, f"Unable to access config properties for\n\n  '{node_path}'") from e

    @staticmethod
    def get_paths(setting: str | InfoPathsKind) -> str:
        """
        Gets access to LO's predefined paths.

        |lo_unsafe|

        Args:
            setting (str | InfoPathsKind): property value

        Raises:
            ValueError: if unable to get paths

        Returns:
            str: paths

        Note:
            There are two different groups of properties.
            One group stores only a single path and the other group stores two or
            more paths - separated by a semicolon.

            Some setting values (as listed in the OpenOffice docs for PathSettings) are
            found in :py:class:`~.kind.info_paths_kind.InfoPathsKind`

        See Also:
            - :py:class:`~.kind.info_paths_kind.InfoPathsKind`
            - :py:meth:`Info.get_dirs`
            - :ref:`ch03`
            - `Wiki Path Settings <https://wiki.openoffice.org/w/index.php?title=Documentation/DevGuide/OfficeDev/Path_Settings>`_
        """
        # access LO's predefined paths. There are two different groups of properties.
        #  One group stores only a single path and the other group stores two or
        #  more paths - separated by a semicolon. See
        #  https://wiki.openoffice.org/w/index.php?title=Documentation/DevGuide/OfficeDev/Path_Settings

        #  Some setting values (as listed in the OpenOffice docs for PathSettings):
        #    Addin, AutoCorrect, AutoText, Backup, Basic, Bitmap,
        #    Config, Dictionary, Favorite, Filter, Gallery,
        #    Graphic, Help, Linguistic, Module, Palette, Plugin,
        #    Storage, Temp, Template, UIConfig, UserConfig,
        #    UserDictionary (deprecated), Work

        # Replaced by thePathSetting in LibreOffice 4.3
        try:
            prop_set = mLo.Lo.create_instance_mcf(XPropertySet, "com.sun.star.util.PathSettings", raise_err=True)
            result = prop_set.getPropertyValue(str(setting))
            if result is None:
                raise ValueError(f"getPropertyValue() for {setting} yielded None")
            return str(result)
        except Exception as e:
            raise ValueError(f"Could not find paths for: {setting}") from e

    @classmethod
    def get_dirs(cls, setting: str) -> List[Path]:
        """
        Gets dirs paths from settings.

        |lo_unsafe|

        Args:
            setting (str): setting

        Returns:
            List[str]: List of paths

        See Also:
            :py:meth:`~.Info.get_paths`

            `Wiki Path Settings <https://wiki.openoffice.org/w/index.php?title=Documentation/DevGuide/OfficeDev/Path_Settings>`_

        .. versionchanged:: 0.11.14
            - Added support for expanding macros.
        """
        try:
            paths = cls.get_paths(setting)
        except ValueError:
            mLo.Lo.print(f"Could not find paths for '{setting}'")
            return []
        paths_arr = paths.split(";")
        if len(paths_arr) == 0:
            mLo.Lo.print(f"Could not split paths for '{setting}'")
            return [mFileIO.FileIO.uri_to_path(mFileIO.FileIO.uri_absolute(mFileIO.FileIO.expand_macro(paths)))]
        return [
            mFileIO.FileIO.uri_to_path(mFileIO.FileIO.uri_absolute(mFileIO.FileIO.expand_macro(el)))
            for el in paths_arr
        ]

    @classmethod
    def get_office_dir(cls) -> str:
        """
        Gets file path to the office dir.
        e.g. ``"C:\\Program Files (x86)\\LibreOffice 7"``.

        |lo_unsafe|

        Raises:
            ValueError: if unable to obtain office path.

        Returns:
            str: Path as string

        See Also:
            :ref:`ch03`
        """
        try:
            addin_dir = cls.get_paths("Addin")
            # e.g. 'file:///C:/Program%20Files/LibreOffice/program/../program/addin

            addin_path = str(mFileIO.FileIO.uri_to_path(addin_dir))
            #   e.g.  C:\Program%20Files\LibreOffice\program\addin
            try:
                idx = addin_path.index("program")
            except ValueError:
                mLo.Lo.print("Could not extract office path")
                return addin_path

            p = Path(addin_path[:idx])
            # e.g.  'C:\Program%20Files\LibreOffice\
            return str(p)
        except Exception as e:
            raise ValueError("Unable to get office dir") from e

    @classmethod
    def get_office_theme(cls) -> str:
        """
        Gets the Current LibreOffice Theme.

        |lo_unsafe|

        Returns:
            str: LibreOffice Theme Name such as ``LibreOffice Dark``

        Note:
            Older Version of LibreOffice will return empty string.

        .. versionadded:: 0.9.1
        """
        try:
            return str(
                cls.get_config(
                    node_str="CurrentColorScheme",
                    node_path="/org.openoffice.Office.ExtendedColorScheme/ExtendedColorScheme",
                )
            )
        except mEx.ConfigError:
            # most likely pre LO 7.4
            return ""

    @classmethod
    def get_gallery_dir(cls) -> Path:
        """
        Get the first directory that contain the Gallery database and multimedia files.

        |lo_unsafe|

        Raises:
            ValueError if unable to obtain gallery dir.

        Returns:
            Path: Gallery Dir

        See Also:
            :ref:`ch03`
        """
        try:
            gallery_dirs = cls.get_dirs("Gallery")
            if gallery_dirs is None:
                raise ValueError("No result from get_dir for Gallery")
            return gallery_dirs[0]
        except Exception as e:
            raise ValueError("Unable to get gallery dir") from e

    @classmethod
    def create_configuration_view(cls, path: str) -> XHierarchicalPropertySet:
        """
        Create Configuration View.

        |lo_unsafe|

        Args:
            path (str): path

        Raises:
            ConfigError: if unable to create configuration view

        Returns:
            XHierarchicalPropertySet: Property Set
        """
        try:
            con_prov = mLo.Lo.create_instance_mcf(
                XMultiServiceFactory, "com.sun.star.configuration.ConfigurationProvider"
            )
            if con_prov is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            _props = mProps.Props.make_props(nodepath=path)
            root = con_prov.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", _props)
            # cls.show_services(obj_name="ConfigurationAccess", obj=root)
            ps = mLo.Lo.qi(XHierarchicalPropertySet, root)
            if ps is None:
                raise mEx.MissingInterfaceError(XHierarchicalPropertySet)
            return ps
        except Exception as e:
            raise mEx.ConfigError(f"Unable to get configuration view for '{path}'") from e

    # =================== update configuration settings ================

    @staticmethod
    def set_config_props(node_path: str) -> XPropertySet:
        """
        Get config properties.

        |lo_unsafe|

        Args:
            node_path (str): Node path of properties

        Raises:
            mEx.ConfigError: if Unable to get config properties

        Returns:
            XPropertySet: Property Set
        """
        try:
            con_prov = mLo.Lo.create_instance_mcf(
                XMultiServiceFactory, "com.sun.star.configuration.ConfigurationProvider"
            )
            if con_prov is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            _props = mProps.Props.make_props(nodepath=node_path)
            ca = con_prov.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", _props)
            ps = mLo.Lo.qi(XPropertySet, ca)
            if ps is None:
                raise mEx.MissingInterfaceError(XPropertySet)
            return ps
        except Exception as e:
            raise mEx.ConfigError(f"Unable to set configuration property for '{node_path}'") from e

    @classmethod
    def set_config(cls, node_path: str, node_str: str, val: object) -> bool:
        """
        Sets config.

        |lo_unsafe|

        Args:
            node_path (str): node path
            node_str (str): node name
            val (object): node value

        Returns:
            bool: True on success; Otherwise, False
        """
        with contextlib.suppress(Exception):
            props = cls.set_config_props(node_path=node_path)
            if props is None:
                return False
            mProps.Props.set_property(props, node_str, val)
            secure_change = mLo.Lo.qi(XChangesBatch, props)
            if secure_change is None:
                raise mEx.MissingInterfaceError(XChangesBatch)
            secure_change.commitChanges()
            return True
        return False

    # =================== getting info about a document ====================

    @staticmethod
    def get_name(fnm: PathOrStr) -> str:
        """
        Gets the file's name from the supplied string minus the extension.

        |lo_safe|

        Args:
            fnm (PathOrStr): File path

        Raises:
            ValueError: If fnm is empty string
            ValueError: If fnm is not a file

        Returns:
            str: File name minus the extension
        """
        if fnm == "":
            raise ValueError("Empty string")
        p = Path(fnm)
        if not p.is_file():
            raise ValueError(f"Not a file: '{fnm}'")
        if not p.suffix:
            mLo.Lo.print(f"No extension found for '{fnm}'")
            return p.stem
        return p.stem

    @staticmethod
    def get_ext(fnm: PathOrStr) -> str | None:
        """
        Gets file extension without the ``.``.

        |lo_safe|

        Args:
            fnm (PathOrStr): file path

        Raises:
            ValueError: If fnm is empty string

        Returns:
            str | None: Extension if Found; Otherwise, None
        """
        return mFileIO.FileIO.get_ext(fnm)

    @staticmethod
    def get_unique_fnm(fnm: PathOrStr) -> str:
        """
        If a file called fnm already exists, then a number
        is added to the name so the filename is unique.

        |lo_safe|

        Args:
            fnm (str): file path

        Returns:
            str: unique file path
        """
        p = Path(fnm)
        fname = p.stem
        ext = p.suffix
        i = 1
        while p.exists():
            name = f"{fname}{i}{ext}"
            p = p.parent / name
            i += 1
        return str(p)

    @staticmethod
    def get_doc_type(fnm: PathOrStr) -> str:
        """
        Gets doc type from file path.

        |lo_unsafe|

        Args:
            fnm (PathOrStr): File Path

        Raises:
            ValueError: if Unable to get doc type

        Returns:
            str: Doc Type.
        """
        try:
            x_detect = mLo.Lo.create_instance_mcf(
                XTypeDetection, "com.sun.star.document.TypeDetection", raise_err=True
            )
            if not mFileIO.FileIO.is_openable(fnm):
                raise mEx.UnOpenableError(fnm)
            url_str = str(mFileIO.FileIO.fnm_to_url(fnm))
            media_desc = (mProps.Props.make_prop_value(name="URL", value=url_str),)

            # even thought queryTypeByDescriptor reports to return a string
            # for some reason I am getting a tuple with the first value as the expected result.
            result = x_detect.queryTypeByDescriptor(media_desc, True)
            if result is None:
                raise mEx.UnKnownError("queryTypeByDescriptor() is an unknown result")
            return result if isinstance(result, str) else result[0]
        except Exception as e:
            raise ValueError(f"unable to get doc type for '{fnm}'") from e

    @classmethod
    def report_doc_type(cls, doc: Any) -> mLo.Lo.DocType:
        """
        Prints doc type to console and return doc type.

        |lo_safe|

        Args:
            doc (object): office document

        Returns:
            Lo.DocType: Document type.
        """
        doc_type = mLo.Lo.DocType.UNKNOWN
        if cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.WRITER):
            mLo.Lo.print("A Writer document")
            doc_type = mLo.Lo.DocType.WRITER
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.IMPRESS):
            mLo.Lo.print("A Impress document")
            doc_type = mLo.Lo.DocType.IMPRESS
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.DRAW):
            mLo.Lo.print("A Draw document")
            doc_type = mLo.Lo.DocType.DRAW
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.CALC):
            mLo.Lo.print("A Calc document")
            doc_type = mLo.Lo.DocType.CALC
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.BASE):
            mLo.Lo.print("A Base document")
            doc_type = mLo.Lo.DocType.BASE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.MATH):
            mLo.Lo.print("A Math document")
            doc_type = mLo.Lo.DocType.MATH
        else:
            mLo.Lo.print("Unknown document")
        return doc_type

    @classmethod
    def doc_type_service(cls, doc: Any) -> mLo.Lo.Service:
        """
        Prints service type to console and return service type.

        |lo_safe|

        Args:
            doc (Any): office document

        Returns:
            Lo.Service: Service type
        """
        if cls.is_doc_type(obj=doc, doc_type=LoService.WRITER):
            mLo.Lo.print("A Writer document")
            return mLo.Lo.Service.WRITER
        elif cls.is_doc_type(obj=doc, doc_type=LoService.IMPRESS):
            mLo.Lo.print("A Impress document")
            return mLo.Lo.Service.IMPRESS
        elif cls.is_doc_type(obj=doc, doc_type=LoService.DRAW):
            mLo.Lo.print("A Draw document")
            return mLo.Lo.Service.DRAW
        elif cls.is_doc_type(obj=doc, doc_type=LoService.CALC):
            mLo.Lo.print("A Calc document")
            return mLo.Lo.Service.CALC
        elif cls.is_doc_type(obj=doc, doc_type=LoService.BASE):
            mLo.Lo.print("A Base document")
            return mLo.Lo.Service.BASE
        elif cls.is_doc_type(obj=doc, doc_type=LoService.MATH):
            mLo.Lo.print("A Math document")
            return mLo.Lo.Service.MATH
        else:
            mLo.Lo.print("Unknown document")
            return mLo.Lo.Service.UNKNOWN

    @staticmethod
    def is_doc_type(obj: Any, doc_type: LoService | str) -> bool:
        """
        Gets if doc is a particular doc type.

        |lo_safe|

        Args:
            obj (object): office document
            doc_type (Service, str): Doc type or service name such as ``com.sun.star.text.TextDocument``.

        Returns:
            bool: ``True`` if obj matches; Otherwise, ``False``
        """
        with contextlib.suppress(Exception):
            si = mLo.Lo.qi(XServiceInfo, obj)
            return False if si is None else si.supportsService(str(doc_type))
        return False

    @staticmethod
    def get_implementation_name(obj: Any) -> str:
        """
        Gets implementation name such as ``com.sun.star.comp.deployment.PackageInformationProvider``.

        |lo_safe|

        Args:
            obj (object): uno object that implements XServiceInfo

        Raises:
            ValueError: if unable to get implementation name

        Returns:
            str: implementation name
        """
        try:
            si = mLo.Lo.qi(XServiceInfo, obj, True)
            return si.getImplementationName()
        except Exception as e:
            raise ValueError("Could not get service information") from e

    @staticmethod
    def get_identifier(obj: Any) -> str:
        """
        Gets identifier name such as ``com.sun.star.text.TextDocument``.

        |lo_safe|

        Args:
            obj (object): uno object that implements XModule

        Raises:
            ValueError: if unable to get identifier name

        Returns:
            str: identifier name
        """
        try:
            x_mod = mLo.Lo.qi(XModule, obj, True)
            return x_mod.getIdentifier()
        except Exception as e:
            raise ValueError("Could not get service information") from e

    @staticmethod
    def get_mime_type(fnm: PathOrStr) -> str:
        """
        Get mime type for a file path.

        |lo_safe|

        Args:
            fnm (PathOrStr): file path

        Returns:
            str: Mime type of file if found. Defaults to 'application/octet-stream'
        """
        default = "application/octet-stream"
        mt = mimetypes.guess_type(fnm)
        if mt is None:
            mLo.Lo.print("unable to find mimetype")
            return default
        if mt[0] is None:
            mLo.Lo.print("unable to find mimetype")
            return default
        return str(mt[0])

    @staticmethod
    def mime_doc_type(mime_type: str) -> mLo.Lo.DocType:
        """
        Gets document type from mime type.

        |lo_safe|

        Args:
            mime_type (str): mime type

        Returns:
            Lo.DocType: Document type. If mime_type is unknown then 'DocType.UNKNOWN'
        """
        if mime_type is None or not mime_type:
            return mLo.Lo.DocType.UNKNOWN
        if "vnd.oasis.opendocument.text" in mime_type:
            return mLo.Lo.DocType.WRITER
        if "vnd.oasis.opendocument.base" in mime_type:
            return mLo.Lo.DocType.BASE
        if "vnd.oasis.opendocument.spreadsheet" in mime_type:
            return mLo.Lo.DocType.CALC
        if (
            "vnd.oasis.opendocument.graphics" in mime_type
            or "vnd.oasis.opendocument.image" in mime_type
            or "vnd.oasis.opendocument.chart" in mime_type
        ):
            return mLo.Lo.DocType.DRAW
        if "vnd.oasis.opendocument.presentation" in mime_type:
            return mLo.Lo.DocType.IMPRESS
        if "vnd.oasis.opendocument.formula" in mime_type:
            return mLo.Lo.DocType.MATH
        return mLo.Lo.DocType.UNKNOWN

    @staticmethod
    def is_image_mime(mime_type: str) -> bool:
        """
        Gets if mime-type is a known image type.

        |lo_safe|

        Args:
            mime_type (str): mime type e.g. ``application/x-openoffice-bitmap``

        Returns:
            bool: True if known mime-type; Otherwise False
        """
        if mime_type is None or not mime_type:
            return False
        if mime_type.startswith("image/"):
            return True
        return bool(mime_type.startswith("application/x-openoffice-bitmap"))

    # ------------------------ services, interfaces, methods info ----------------------
    @overload
    @classmethod
    def get_service_names(cls) -> List[str]:
        """
        Gets service names.

        |lo_unsafe|

        Returns:
            List[str]: List of service names.
        """
        ...

    @overload
    @classmethod
    def get_service_names(cls, service_name: str) -> List[str]:
        """
        Gets service names.

        |lo_unsafe|

        Args:
            service_name (str): service name.

        Returns:
            List[str]: List of service names.
        """
        ...

    @classmethod
    def get_service_names(cls, service_name: Optional[str] = None) -> List[str]:
        """
        Gets service names.

        |lo_unsafe|

        Args:
            service_name (str): service name.

        Raises:
            Exception: If error occurs.

        Returns:
            List[str]: List of service names.
        """
        if service_name is None:
            return cls._get_service_names1()
        return cls._get_service_names2(service_name=service_name)

    @staticmethod
    def _get_service_names1() -> List[str]:
        """LO UN-safe method"""
        mc_factory = mLo.Lo.get_component_factory()
        if mc_factory is None:
            return []
        return sorted(mc_factory.getAvailableServiceNames())

    @staticmethod
    def _get_service_names2(service_name: str) -> List[str]:
        """LO Un-safe method"""
        names: List[str] = []
        try:
            enum_access = mLo.Lo.qi(XContentEnumerationAccess, mLo.Lo.get_component_factory(), True)
            x_enum = enum_access.createContentEnumeration(service_name)
            while x_enum.hasMoreElements():
                si = mLo.Lo.qi(XServiceInfo, x_enum.nextElement(), True)
                names.append(si.getImplementationName())
        except Exception as e:
            mLo.Lo.print(f"Could not collect service names for: {service_name}")
            raise e
        if not names:
            mLo.Lo.print(f"No service names found for: {service_name}")
            return names

        names.sort()
        return names

    @staticmethod
    def get_services(obj: Any) -> List[str]:
        """
        Gets service names.

        |lo_safe|

        Args:
            obj (object): obj that implements XServiceInfo

        Returns:
            List[str]: service names
        """
        with contextlib.suppress(AttributeError):
            return sorted(obj.SupportedServiceNames)  # type: ignore
        try:
            si = mLo.Lo.qi(XServiceInfo, obj, True)
            names = si.getSupportedServiceNames()
            return sorted(names)
        except Exception as e:
            mLo.Lo.print("Unable to get services")
            mLo.Lo.print(f"    {e}")
            raise e

    @classmethod
    def show_services(cls, obj_name: str, obj: Any) -> None:
        """
        Prints services to console.

        |lo_safe|

        Args:
            obj_name (str): service name
            obj (object): obj that implements XServiceInfo
        """
        services = cls.get_services(obj=obj)
        if services is None:
            print(f"No supported services found for {obj_name}")
            return
        print(f"{obj_name} Supported Services ({len(services)})")
        for service in services:
            print(f"'{service}'")

    @staticmethod
    def support_service(obj: Any, *service: str) -> bool:
        """
        Gets if ``obj`` supports a service.

        |lo_safe|

        Args:
            obj (object): Object to check for supported service
            *service (str): Variable length argument list of UNO namespace strings such as ``com.sun.star.configuration.GroupAccess``

        Returns:
            bool: ``True`` if ``obj`` supports any passed in service; Otherwise, ``False``
        """

        result = False
        try:
            si = mLo.Lo.qi(XServiceInfo, obj)
            if si is None:
                return result
            for srv in service:
                result = si.supportsService(srv)  # type: ignore
                if result:
                    break
        except Exception as e:
            mLo.Lo.print("Errors ocurred in support_service(). Returning False")
            mLo.Lo.print(f"    {e}")
        return result

    @staticmethod
    def get_available_services(obj: Any) -> List[str]:
        """
        Gets available services for obj.

        |lo_safe|

        Args:
            obj (object): obj that implements ``XMultiServiceFactory`` interface

        Raises:
            Exception: If unable to get services

        Returns:
            List[str]: List of services
        """
        # sourcery skip: raise-specific-error
        services: List[str] = []
        try:
            sf = mLo.Lo.qi(XMultiServiceFactory, obj, True)
            service_names = sf.getAvailableServiceNames()
            services.extend(service_names)
            services.sort()
        except Exception as e:
            mLo.Lo.print(e)
            raise Exception() from e
        return services

    @staticmethod
    def get_interface_types(target: object) -> Tuple[object, ...]:
        """
        Get interface types.

        |lo_safe|

        Args:
            target (object): object that implements XTypeProvider interface.

        Raises:
            Exception: If unable to get services.

        Returns:
            Tuple[object, ...]: Tuple of interfaces.
        """
        # sourcery skip: raise-specific-error
        try:
            tp = mLo.Lo.qi(XTypeProvider, target, True)
            return tp.getTypes()
        except Exception as e:
            mLo.Lo.print("Unable to get interface types")
            mLo.Lo.print(f"    {e}")
            raise Exception() from e

    @overload
    @classmethod
    def get_interfaces(cls, target: object) -> List[str]:
        """
        Gets interfaces.

        |lo_safe|

        Args:
            target: (object): object that implements XTypeProvider.

        Returns:
            List[str]: List of interfaces.
        """
        ...

    @overload
    @classmethod
    def get_interfaces(cls, type_provider: XTypeProvider) -> List[str]:
        """
        Gets interfaces.

        |lo_safe|

        Args:
            type_provider (XTypeProvider): type provider.

        Returns:
            List[str]: List of interfaces
        """
        ...

    @classmethod
    def get_interfaces(cls, *args, **kwargs) -> List[str]:
        """
        Gets interfaces.

        |lo_safe|

        Args:
            target: (object): object that implements ``XTypeProvider``.
            type_provider (XTypeProvider): type provider.

        Raises:
            Exception: If unable to get interfaces.

        Returns:
            List[str]: List of interfaces.
        """
        # sourcery skip: raise-specific-error
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("target", "typeProvider")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_interfaces() got an unexpected keyword argument")
            keys = ("target", "typeProvider")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("get_interfaces() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        try:
            if mLo.Lo.is_uno_interfaces(kargs[1], XTypeProvider):
                type_provider = cast(XTypeProvider, kargs[1])
            else:
                type_provider = mLo.Lo.qi(XTypeProvider, kargs[1])
                if type_provider is None:
                    raise mEx.MissingInterfaceError(XTypeProvider)

            types = cast(Tuple[uno.Type, ...], type_provider.getTypes())
            # use a set to exclude duplicate names
            names_set: Set[str] = set()
            for t in types:
                names_set.add(t.typeName)
            type_names = sorted(list(names_set))
            return type_names
        except Exception as e:
            mLo.Lo.print("Unable to get interfaces")
            mLo.Lo.print(f"    {e}")
            raise Exception() from e

    @classmethod
    def is_interface_obj(cls, obj: Any) -> bool:
        """
        Gets is an object contains interfaces.

        |lo_safe|

        Args:
            obj (Any): Object to check.

        Returns:
            bool: ``True`` if obj contains interface; Otherwise, ``False``.
        """
        result = False
        if obj is None:
            return result
        with contextlib.suppress(mEx.MissingInterfaceError):
            interfaces = cls.get_interfaces(obj)
            result = len(interfaces) > 0
        return result

    @staticmethod
    def is_struct(obj: Any) -> bool:
        """
        Gets if an object is a UNO Struct.

        |lo_safe|

        Args:
            obj (Any): Object to check.

        Returns:
            bool: ``True`` if obj is Struct; Otherwise, ``False``
        """
        if obj is None:
            return False
        if not hasattr(obj, "typeName"):
            return False
        with contextlib.suppress(Exception):
            t = cast(uno_proto.UnoType, uno.getTypeByName(obj.typeName))
            return t.typeClass.value == "STRUCT"
        return False

    @classmethod
    def is_same(cls, obj1: Any, obj2: Any) -> bool:
        """
        Determines if two Uno object are the same.

        |lo_safe|

        Args:
            obj1 (Any): First Uno object
            obj2 (Any): Second Uno Object

        Returns:
            bool: True if objects are the same; Otherwise False.
        """
        # https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/XInterface
        # In C++, two objects are the same if their XInterface are the same. The queryInterface() for XInterface will have to
        # be called on both. In Java, check for the identity by calling the runtime function
        # com.sun.star.uni.UnoRuntime.areSame().
        if cls.is_struct(obj1) and cls.is_struct(obj2):
            # types and attribute values must match
            if obj1.typeName != obj2.typeName:
                return False
            obj1_attrs = [s for s in dir(obj1.value) if not s.startswith("_")]
            return all(getattr(obj1, atr) == getattr(obj2, atr) for atr in obj1_attrs)
        elif cls.is_interface_obj(obj1) and cls.is_interface_obj(obj2):
            # must be same object in memory
            id1 = id(obj1)
            id2 = id(obj2)
            return id1 == id2
        return obj1 == obj2

    @staticmethod
    def show_container_names(print_name: str, nc: XNameContainer):
        """
        Prints Name Container elements to console.

        |lo_safe|

        Args:
            print_name (str): Name to display. Can be empty string.
            nc (XNameContainer): Name Container to print element names of.
        """
        names = nc.getElementNames()
        if print_name:
            print(f"No. of Names in {print_name}: {len(names)}")
        else:
            print(f"No. of Names in Name Container: {len(names)}")
        for name in names:
            print(f"  {name}")
        print()

    @classmethod
    def show_interfaces(cls, obj_name: str, obj: Any) -> None:
        """
        Prints interfaces in obj to console.

        |lo_safe|

        Args:
            obj_name (str): Name of object for printing
            obj (object): obj that contains interfaces.
        """
        interfaces = cls.get_interfaces(obj)
        if not interfaces:
            print(f"No interfaces found for {obj_name}")
            return
        print(f"{obj_name} Interfaces ({len(interfaces)})")
        for s in interfaces:
            print(f"  {s}")

    @staticmethod
    def show_conversion_values(value: Any, frm: mConvert.UnitLength) -> None:
        """
        Prints values of conversions to terminal.

        |lo_safe|

        Args:
            value (Any): Any numeric value
            frm (mConvert.Length): The unit type of ``value``.

        Note:
            Useful to help determine what different conversion of ``value`` are.

        .. versionadded:: 0.9.0
        """
        lengths = cast(
            List[mConvert.UnitLength],
            [
                getattr(mConvert.UnitLength, x)
                for x in dir(mConvert.UnitLength)
                if x.isupper()
                and getattr(mConvert.UnitLength, x).value < mConvert.UnitLength.COUNT
                and getattr(mConvert.UnitLength, x).value >= 0
            ],
        )
        for length in lengths:
            try:
                result = mConvert.UnitConvert.convert(num=value, frm=frm, to=length)
                print(f"{length.name}:".ljust(10), round(result, 2))
            except Exception:
                print(f'"{length.name}" does not convert')

    @staticmethod
    def get_methods_obj(obj: Any, property_concept: PropertyConceptEnum | None = None) -> List[str]:
        """
        Get Methods of an object such as a doc.

        |lo_safe|

        Args:
            obj (object): Object to get methods of.
            property_concept (PropertyConceptEnum | None, optional): Type of method to get. Defaults to PropertyConceptEnum.ALL.

        Raises:
            Exception: If unable to get Methods

        Returns:
            List[str]: List of method names found for obj
        """
        # sourcery skip: raise-specific-error
        if property_concept is None:
            property_concept = PropertyConceptEnum.ALL

        try:
            intro = theIntrospection()  # type: ignore
            result = intro.inspect(obj)
            methods = result.getMethods(int(property_concept))
            lst = [meth.getName() for meth in methods]
            lst.sort()
            return lst
        except Exception as e:
            raise Exception("Could not get object Methods") from e

    @staticmethod
    def get_methods(interface_name: str) -> List[str]:
        """
        Get Interface Methods.

        |lo_unsafe|

        Args:
            interface_name (str): name of interface.

        Returns:
            List[str]: List of methods
        """
        # sourcery skip: raise-specific-error
        # from com.sun.star.beans.PropertyConcept import ALL
        # ctx = XSCRIPTCONTEXT.getComponentContext()
        # smgr = ctx.getServiceManager()
        # intro = smgr.createInstance('com.sun.star.beans.Introspection')
        # document = XSCRIPTCONTEXT.getDocument()
        # result = intro.inspect(document)
        # methods = result.getMethods(ALL)
        # meth = methods[0]
        #   meth is com.sun.star.reflection.XIdlMethod
        # meth.Name
        #   queryInterface
        #
        #  reflection = smgr.createInstanceWithContext('com.sun.star.reflection.CoreReflection', ctx)
        # See Also: https://tinyurl.com/y338vwhw#L492
        # See Also: https://github.com/hanya/MRI/wiki/RunMRI#Python
        # See Also: https://tinyurl.com/y3m4tx9r#L268

        reflection = mLo.Lo.create_instance_mcf(
            XIdlReflection, "com.sun.star.reflection.CoreReflection", raise_err=True
        )

        fname = reflection.forName(interface_name)  # returns type from name.

        if fname is None:
            mLo.Lo.print(f"Could not find the interface name: {interface_name}")
            return []
        try:
            methods: Tuple[XIdlMethod, ...] = fname.getMethods()
            lst = [meth.getName() for meth in methods]
            lst.sort()
            return lst
        except Exception as e:
            raise Exception(f"Could not get Methods for: {interface_name}") from e

    @classmethod
    def show_methods(cls, interface_name: str) -> None:
        """
        Prints methods to console for an interface.

        |lo_unsafe|

        Args:
            interface_name (str): name of interface.
        """
        methods = cls.get_methods(interface_name=interface_name)
        if len(methods) == 0:
            print(f"No methods found for '{interface_name}'")
            return
        print(f"{interface_name} Methods: {len(methods)}")
        for method in methods:
            print(f"  {method}")

    @classmethod
    def show_methods_obj(cls, obj: Any, property_concept: PropertyConceptEnum | None = None) -> None:
        """
        Prints method to console for an object such as a doc.

        |lo_safe|

        Args:
            obj (object): Object to get methods of.
            property_concept (PropertyConceptEnum | None, optional): Type of method to get. Defaults to PropertyConceptEnum.ALL.
        """
        methods = cls.get_methods_obj(obj=obj, property_concept=property_concept)
        if len(methods) == 0:
            print("No Object methods found")
            return
        print(f"Object Methods: {len(methods)}")
        for method in methods:
            print(f"  {method}")

    # -------------------------- style info --------------------------
    @staticmethod
    def get_style_families(doc: Any) -> XNameAccess:
        """
        Gets a list of style family names.

        |lo_safe|

        Args:
            doc (Any): office document.

        Raises:
            MissingInterfaceError: If Doc does not implement XStyleFamiliesSupplier interface.
            Exception: If unable to get style Families.

        Returns:
            XNameAccess: Style Families.
        """
        # sourcery skip: raise-specific-error
        try:
            xsupplier = mLo.Lo.qi(XStyleFamiliesSupplier, doc)
            if xsupplier is None:
                raise mEx.MissingInterfaceError(XStyleFamiliesSupplier)
            return xsupplier.getStyleFamilies()
        except mEx.MissingInterfaceError:
            raise
        except Exception as e:
            raise Exception("Unable to get family style names") from e

    @classmethod
    def get_style_family_names(cls, doc: Any) -> List[str]:
        """
        Gets a sorted list of style family names.

        |lo_safe|

        Args:
            doc (Any): office document

        Raises:
            Exception: If unable to names.

        Returns:
            List[str]: List of style names
        """
        # sourcery skip: raise-specific-error
        try:
            name_acc = cls.get_style_families(doc)
            names = name_acc.getElementNames()
            return sorted(names)
        except Exception as e:
            raise Exception("Unable to get family style names") from e

    @classmethod
    def get_style_container(cls, doc: Any, family_style_name: str) -> XNameContainer:
        """
        Gets style container of document for a family of styles.

        |lo_safe|

        Args:
            doc (Any): office document
            family_style_name (str): Family style name

        Raises:
            MissingInterfaceError: if unable to obtain XNameContainer interface

        Returns:
            XNameContainer: Style Family container
        """
        name_acc = cls.get_style_families(doc)
        return mLo.Lo.qi(XNameContainer, name_acc.getByName(family_style_name), raise_err=True)

    @classmethod
    def get_style_names(cls, doc: Any, family_style_name: str) -> List[str]:
        """
        Gets a list of style names.

        |lo_safe|

        Args:
            doc (Any): office document.
            family_style_name (str): name of family style.

        Raises:
            Exception: If unable to access Style names.

        Returns:
            List[str]: List of style names.
        """
        # sourcery skip: raise-specific-error
        try:
            style_container = cls.get_style_container(doc=doc, family_style_name=family_style_name)
            names = style_container.getElementNames()
            return sorted(names)
        except Exception as e:
            raise Exception("Could not access style names") from e

    @classmethod
    def get_style_props(cls, doc: Any, family_style_name: str, prop_set_nm: str) -> XPropertySet:
        """
        Get style properties for a family of styles.

        |lo_safe|

        Args:
            doc (Any): office document.
            family_style_name (str): name of family style.
            prop_set_nm (str): property set name.

        Raises:
            MissingInterfaceError: if a required interface cannot be obtained.

        Returns:
            XPropertySet: Property set.
        """
        style_container = cls.get_style_container(doc, family_style_name)
        return mLo.Lo.qi(XPropertySet, style_container.getByName(prop_set_nm), True)

    @classmethod
    def get_page_style_props(cls, doc: Any) -> XPropertySet:
        """
        Gets style properties for page styles.

        |lo_safe|

        Args:
            doc (Any): office docs.

        Raises:
            MissingInterfaceError: if a required interface cannot be obtained.

        Returns:
            XPropertySet: property set.
        """
        return cls.get_style_props(doc, "PageStyles", "Standard")

    @classmethod
    def get_paragraph_style_props(cls, doc: Any) -> XPropertySet:
        """
        Gets style properties for paragraph styles.

        |lo_safe|

        Args:
            doc (Any): office docs.

        Raises:
            MissingInterfaceError: if a required interface cannot be obtained.

        Returns:
            XPropertySet: property set.
        """
        return cls.get_style_props(doc, "ParagraphStyles", "Standard")

    # ----------------------------- document properties ----------------------

    @classmethod
    def print_doc_properties(cls, doc: Any) -> None:
        """
        Prints document properties to console.

        |lo_safe|

        Args:
            doc (Any): office document
        """
        try:
            doc_props_supp = mLo.Lo.qi(XDocumentPropertiesSupplier, doc, True)
            dps = doc_props_supp.getDocumentProperties()
            cls.print_doc_props(dps=dps)
            ud_props = dps.getUserDefinedProperties()
            mProps.Props.show_obj_props("UserDefined Info", ud_props)
        except Exception as e:
            mLo.Lo.print("Unable to get doc properties")
            mLo.Lo.print(f"    {e}")
        return

    @classmethod
    def print_doc_props(cls, dps: XDocumentProperties) -> None:
        """
        Prints doc properties to console.

        |lo_safe|

        Args:
            dps (XDocumentProperties): document properties.

        See Also:
            :py:meth:`~Info.print_doc_properties`
        """
        print("Document Properties Info")
        print(f"  Author: {dps.Author}")
        print(f"  Title: {dps.Title}")
        print(f"  Subject: {dps.Subject}")
        print(f"  Description: {dps.Description}")
        print(f"  Generator: {dps.Generator}")

        keys = dps.Keywords
        print("  Keywords: ")
        for keyword in keys:
            print(f"  {keyword}")

        print(f"  Modified by: {dps.ModifiedBy}")
        print(f"  Printed by: {dps.PrintedBy}")
        print(f"  Template Name: {dps.TemplateName}")
        print(f"  Template URL: {dps.TemplateURL}")
        print(f"  Autoload URL: {dps.AutoloadURL}")
        print(f"  Default Target: {dps.DefaultTarget}")

        l = dps.Language
        loc = [
            "unknown" if len(l.Language) == 0 else l.Language,
            "unknown" if len(l.Country) == 0 else l.Country,
            "unknown" if len(l.Variant) == 0 else l.Variant,
        ]
        print(f"  Locale: {'; '.join(loc)}")

        print(f"  Modification Date: {mDate.DateUtil.str_date_time(dps.ModificationDate)}")
        print(f"  Creation Date: {mDate.DateUtil.str_date_time(dps.CreationDate)}")
        print(f"  Print Date: {mDate.DateUtil.str_date_time(dps.PrintDate)}")
        print(f"  Template Date: {mDate.DateUtil.str_date_time(dps.TemplateDate)}")

        doc_stats = dps.DocumentStatistics
        print("  Document statistics:")
        for nv in doc_stats:
            print(f"  {nv.Name} = {nv.Value}")

        try:
            print(f"  Autoload Secs: {dps.AutoloadSecs}")
        except Exception as e:
            print(f"  Autoload Secs: {e}")
        try:
            print(f"  Editing Cycles: {dps.EditingCycles}")
        except Exception as e:
            print(f"  Editing Cycles: {e}")
        try:
            print(f"  Editing Duration: {dps.EditingDuration}")
        except Exception as e:
            print(f"  Editing Duration: {e}")
        print()

    @staticmethod
    def set_doc_props(doc: Any, subject: str, title: str, author: str) -> None:
        """
        Set document properties for subject, title, author.

        |lo_safe|

        Args:
            doc (Any): office document.
            subject (str): subject.
            title (str): title.
            author (str): author.

        Raises:
            PropertiesError: If unable to set properties.
        """
        try:
            dp_supplier = mLo.Lo.qi(XDocumentPropertiesSupplier, doc, True)
            doc_props = dp_supplier.getDocumentProperties()
            doc_props.Subject = subject
            doc_props.Title = title
            doc_props.Author = author
        except Exception as e:
            raise mEx.PropertiesError("Unable to set doc properties") from e

    @staticmethod
    def get_user_defined_props(doc: Any) -> XPropertyContainer:
        """
        Gets user defined properties.

        |lo_safe|

        Args:
            doc (Any): office document.

        Raises:
            PropertiesError: if unable to access properties.

        Returns:
            XPropertyContainer: Property container.

        See Also:
            - :ref:`help_common_modules_info_get_user_defined_props`.
        """
        try:
            dp_supplier = mLo.Lo.qi(XDocumentPropertiesSupplier, doc, True)
            dps = dp_supplier.getDocumentProperties()
            return dps.getUserDefinedProperties()
        except Exception as e:
            raise mEx.PropertiesError("Unable to get user defined props") from e

    # ----------- installed package info -----------------

    @staticmethod
    def get_pip() -> XPackageInformationProvider:
        """
        Gets Package Information Provider.

        |lo_unsafe|

        Raises:
            MissingInterfaceError: if unable to obtain XPackageInformationProvider interface.

        Returns:
            XPackageInformationProvider: Package Information Provider.
        """
        ctx = mLo.Lo.get_context()
        return mLo.Lo.qi(
            XPackageInformationProvider,
            ctx.getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider"),
            True,
        )
        # return pip.get(mLo.Lo.get_context())

    @classmethod
    def list_extensions(cls) -> None:
        """
        Prints extensions to console.

        |lo_unsafe|
        """
        try:
            pip = cls.get_pip()
        except mEx.MissingInterfaceError:
            print("No package info provider found")
            return
        exts_tbl = pip.getExtensionList()
        print("\nExtensions:")
        for i in range(len(exts_tbl)):
            print(f"{i+1}. ID: {exts_tbl[i][0]}")
            print(f"   Version: {exts_tbl[i][1]}")
            print(f"   Loc: {pip.getPackageLocation(exts_tbl[i][0])}")
            print()

    @classmethod
    def get_extension_info(cls, id: str) -> Tuple[str, ...]:
        """
        Gets info for an installed extension in LibreOffice.

        |lo_unsafe|

        Args:
            id (str): Extension id.

        Returns:
            Tuple[str, ...]: Extension info.
        """
        try:
            pip = cls.get_pip()
        except mEx.MissingInterfaceError:
            mLo.Lo.print("No package info provider found")
            return ()
        exts_tbl = pip.getExtensionList()
        mLo.Lo.print_table("Extension", exts_tbl)
        for el in exts_tbl:
            if el[0] == id:
                return el

        mLo.Lo.print(f"Extension {id} is not found")
        return ()

    @classmethod
    def get_extension_loc(cls, id: str) -> str | None:
        """
        Gets location for an installed extension in LibreOffice.

        |lo_unsafe|

        Args:
            id (str): Extension id.

        Returns:
            str | None: Extension location on success; Otherwise, None.
        """
        try:
            pip = cls.get_pip()
        except mEx.MissingInterfaceError:
            mLo.Lo.print("No package info provider found")
            return None
        return pip.getPackageLocation(id)

    @staticmethod
    def get_filter_names() -> Tuple[str, ...]:
        """
        Gets filter names.

        |lo_unsafe|

        Returns:
            Tuple[str, ...]: Filter names.
        """
        na = mLo.Lo.create_instance_mcf(XNameAccess, "com.sun.star.document.FilterFactory")
        if na is None:
            mLo.Lo.print("No Filter factory found")
            return ()
        return na.getElementNames()

    @staticmethod
    def get_filter_props(filter_nm: str) -> List[PropertyValue]:
        """
        Gets filter properties.

        |lo_unsafe|

        Args:
            filter_nm (str): Filter Name.

        Returns:
            List[PropertyValue]: List of PropertyValue.
        """
        na = mLo.Lo.create_instance_mcf(XNameAccess, "com.sun.star.document.FilterFactory")
        if na is None:
            mLo.Lo.print("No Filter factory found")
            return []
        result = cast(Tuple[PropertyValue, ...], na.getByName(filter_nm))
        if result is None:
            mLo.Lo.print(f"No props for filter: {filter_nm}")
            return []
        return list(result)

    @classmethod
    def is_import(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.IMPORT`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.IMPORT) == cls.Filter.IMPORT

    @classmethod
    def is_export(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.EXPORT`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.EXPORT) == cls.Filter.EXPORT

    @classmethod
    def is_template(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.TEMPLATE`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.TEMPLATE) == cls.Filter.TEMPLATE

    @classmethod
    def is_internal(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.INTERNAL`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.INTERNAL) == cls.Filter.INTERNAL

    @classmethod
    def is_template_path(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.TEMPLATEPATH`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.TEMPLATEPATH) == cls.Filter.TEMPLATEPATH

    @classmethod
    def is_own(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.OWN`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.OWN) == cls.Filter.OWN

    @classmethod
    def is_alien(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.ALIEN`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.ALIEN) == cls.Filter.ALIEN

    @classmethod
    def is_default(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.DEFAULT`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.DEFAULT) == cls.Filter.DEFAULT

    @classmethod
    def is_support_selection(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.SUPPORTSSELECTION`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.SUPPORTSSELECTION) == cls.Filter.SUPPORTSSELECTION

    @classmethod
    def is_not_in_file_dialog(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.NOTINFILEDIALOG`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: `True` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.NOTINFILEDIALOG) == cls.Filter.NOTINFILEDIALOG

    @classmethod
    def is_not_in_chooser(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.NOTINCHOOSER`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.NOTINCHOOSER) == cls.Filter.NOTINCHOOSER

    @classmethod
    def is_read_only(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.READONLY`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.READONLY) == cls.Filter.READONLY

    @classmethod
    def is_third_party_filter(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.THIRDPARTYFILTER`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, ``False``.
        """
        return (filter_flags & cls.Filter.THIRDPARTYFILTER) == cls.Filter.THIRDPARTYFILTER

    @classmethod
    def is_preferred(cls, filter_flags: Info.Filter) -> bool:
        """
        Gets if filter flags has ``Filter.PREFERRED`` flag set.

        |lo_safe|

        Args:
            filter_flags (Filter): Flags.

        Returns:
            bool: ``True`` if flag is set; Otherwise, `False`.
        """
        return (filter_flags & cls.Filter.PREFERRED) == cls.Filter.PREFERRED

    @staticmethod
    def is_type_struct(obj: Any, type_name: str) -> bool:
        """
        Gets if an object is a Uno Struct of matching type.

        |lo_safe|

        Args:
            obj (Any): Object to test if is struct.
            type_name (str): Type string such as ``com.sun.star.table.CellRangeAddress``.

        Returns:
            bool: ``True`` if ``obj`` is struct and ``obj`` matches ``type_name``; Otherwise, ``False``.
        """
        if obj is None:
            return False
        return obj.typeName == type_name if hasattr(obj, "typeName") else False

    @staticmethod
    def is_type_interface(obj: Any, type_name: str) -> bool:
        """
        Gets if an object is a Uno interface of matching type.

        |lo_safe|

        Args:
            obj (object): Object to test if is interface
            type_name (str): Type string such as 'com.sun.star.uno.XInterface'

        Returns:
            bool: True if 'obj' is interface and 'obj' matches 'type_name'; Otherwise, False
        """
        return mLo.Lo.is_uno_interfaces(obj, type_name)
        # if obj is None:
        #     return False
        # if hasattr(obj, "__pyunointerface__"):
        #     return obj.__pyunointerface__ == type_name
        # elif hasattr(obj, "queryInterface"):
        #     uno_t = uno.getTypeByName(type_name)
        #     q_obj = obj.queryInterface(uno_t)
        #     if q_obj is not None:
        #         return True
        # return False

    @staticmethod
    def is_type_enum(obj: Any, type_name: str) -> bool:
        """
        Gets if an object is a UNO enum of matching type.

        |lo_safe|

        Args:
            obj (Any): Object to test if is uno enum
            type_name (str): Type string such as ``com.sun.star.sheet.GeneralFunction``

        Returns:
            bool: True if ``obj`` is uno enum and ``obj`` matches ``type_name``; Otherwise, False
        """
        if obj is None:
            return False
        return obj.typeName == type_name if hasattr(obj, "typeName") else False

    @staticmethod
    def is_uno(obj: Any) -> bool:
        """
        Gets if an object is a UNO object.

        |lo_safe|

        Args:
            obj (object): Object to check

        Returns:
            bool: ``True`` if is ``UNO`` object; Otherwise, ``False``

        Note:
            This method only check for type ``pyuno``, therefore UNO ``struct`` are not considered UNO objects in this context.

        .. versionadded:: 0.9.0
        """
        with contextlib.suppress(Exception):
            return type(obj).__name__ == "pyuno"
        return False

    @classmethod
    def is_instance(cls, obj: Any, class_or_tuple: ClassInfo[_T]) -> TypeGuard[_T]:
        """
        Return whether an object is an instance of a class or of a subclass thereof.

        UNO object error when used with ``isinstance``.
        This method will return ``False`` if ``obj`` is a UNO object.

        |lo_safe|

        Args:
            obj (Any): Any object. If UNO object then comparison is done by ``Lo.is_uno_interfaces()``;
                Otherwise, built in ``isinstance`` is used.
            class_or_tuple (Any): A tuple, as in ``isinstance(x, (A, B, ...))``, may be given as the target to check against.
                This is equivalent to ``isinstance(x, A)`` or ``isinstance(x, B)`` or ... etc.
                When a UNO object is being compared then this should be a type or tuple of types that are expected by ``Lo.is_uno_interfaces()``

        Returns:
            bool: ``True`` if is instance; Otherwise, ``False``.

        Example:
            .. code-block:: python

                def __delitem__(self, _item: int | str | DrawPage | XDrawPage) -> None:
                    if mInfo.Info.is_instance(_item, int):
                        self.delete_slide(_item)
                    elif mInfo.Info.is_instance(_item, str):
                        slide = super().get_by_name(_item)
                        if slide is None:
                            raise MissingNameError(f"Unable to find slide with name '{_item}'")
                        super().remove(slide)
                    elif mInfo.Info.is_instance(_item, mDrawPage.DrawPage):
                        super().remove(_item.component)
                    elif mInfo.Info.is_instance(_item, XDrawPage):
                        super().remove(_item)
                    else:
                        raise TypeError(f"Unsupported type: {type(_item)}")

        See Also:
            :py:meth:`~ooodev.utils.lo.Lo.is_uno_interfaces`

        .. versionchanged:: 0.17.14
            Add support to check UNO objects.

        .. versionadded:: 0.17.11
        """
        if cls.is_uno(obj):
            with contextlib.suppress(Exception):
                # Lo.is_uno_interfaces will handle not uno types gracefully
                if isinstance(class_or_tuple, tuple):
                    return mLo.Lo.is_uno_interfaces(obj, *class_or_tuple)
                return mLo.Lo.is_uno_interfaces(obj, class_or_tuple)
            return False
        try:
            return isinstance(obj, class_or_tuple)
        except TypeError:
            return False

    # region is_type_enum_multi()
    @overload
    @staticmethod
    def is_type_enum_multi(alt_type: str, enum_type: Type[Enum], enum_val: Enum) -> bool: ...

    @overload
    @staticmethod
    def is_type_enum_multi(alt_type: str, enum_type: Type[Enum], enum_val: str) -> bool: ...

    @overload
    @staticmethod
    def is_type_enum_multi(alt_type: str, enum_type: Type[Enum], enum_val: Enum, arg_name: str) -> bool: ...

    @overload
    @staticmethod
    def is_type_enum_multi(alt_type: str, enum_type: Type[Enum], enum_val: str, arg_name: str) -> bool: ...

    @staticmethod
    def is_type_enum_multi(alt_type: str, enum_type: Type[Enum], enum_val: Enum | str, arg_name: str = "") -> bool:
        """
        Gets if an multiple inheritance enum, such as a ``str, Enum`` is of expected type.

        |lo_safe|

        Args:
            alt_type (str): Alternative Type, In the case of a ``str, Enum`` this would be ``str``
            enum_type (Type[Enum]): Expected Enum Type
            enum_val (Enum): Actual Enum Value
            arg_name (str, optional): Argument name used to pass enum value into method.

        Raises:
            TypeError: If ``arg_name`` is passed into this method and the type check fails.

        Returns:
            bool: ``True`` if ``enum_val`` is valid ``alt_type`` or ``enum_type``; Otherwise, ``False``

        Example:
            .. code-block:: python

                >>> from ooodev.utils.kind import chart2_types as ct

                >>> val = ct.Area3dKind.PERCENT_STACKED_AREA_3D
                >>> print(is_enum_type("str", ct.ChartTemplateBase, val))
                True
                >>> print(is_enum_type("str", ct.ChartTypeNameBase, val))
                False
                >>> print(is_enum_type("str", ct.ChartTypeNameBase, val, "input_enum"))
                TypeError: Parameter "input_enum" must be of type "str" or "ChartTypeNameBase"
        """
        if isinstance(enum_val, str):
            return True
        if type(enum_val).__name__ != alt_type and not isinstance(enum_val, enum_type):
            if arg_name:
                name = enum_type.__name__
                raise TypeError(f'Parameter "{arg_name}" must be of type "{alt_type}" or "{name}"')
            else:
                return False
        return True

    # endregion is_type_enum_multi()

    @classmethod
    def get_type_name(cls, obj: Any) -> str | None:
        """
        Gets type name such as ``com.sun.star.table.TableSortField`` from uno object.

        |lo_safe|

        Args:
            obj (Any): Uno object

        Returns:
            str | None: Full type name if found; Otherwise; None
        """
        if hasattr(obj, "typeName"):
            return obj.typeName
        if hasattr(obj, "__ooo_full_ns__"):
            # ooouno object
            return obj.__ooo_full_ns__
        return obj.__pyunointerface__ if hasattr(obj, "__pyunointerface__") else None

    # parse_language_code
    @classmethod
    def parse_language_code(cls, lang_code: str) -> Locale:
        """
        Parses a language code into a ``Locale`` object.

        |lo_safe|

        Args:
            lang_code (str): Language code such as ``"en-US"``

        Returns:
            Locale: ``Locale`` object

        Raises:
            ValueError: If ``lang_code`` is not valid.
        """
        if not lang_code:
            raise ValueError("lang_code cannot be empty")
        lang_code = lang_code.lower()
        lang = ""
        country = ""
        variant = ""
        if "-" in lang_code:
            parts = lang_code.split("-", maxsplit=2)
            if len(parts) == 2:
                lang, country = parts
                variant = ""
            elif len(parts) == 3:
                lang, country, variant = parts
        else:
            lang = lang_code
            country = ""
            variant = ""
        if len(lang) != 2:
            raise ValueError(f"Invalid language code: {lang_code}")
        if country and len(country) != 2:
            raise ValueError(f"Invalid country code: {lang_code}")
        return Locale(lang, country.upper(), variant)

    @classmethod
    @deprecated("Use parse_language_code")
    def parse_languange_code(cls, lang_code: str) -> Locale:
        """
        Parses a language code into a ``Locale`` object.

        |lo_safe|

        Args:
            lang_code (str): Language code such as ``"en-US"``

        Returns:
            Locale: ``Locale`` object

        Raises:
            ValueError: If ``lang_code`` is not valid.

        .. deprecated:: 0.9.4
            Use :py:meth:`~.info.Info.parse_language_code` instead.
        """
        return cls.parse_language_code(lang_code=lang_code)

    @classproperty
    def language(cls) -> str:
        """
        Gets the Current Language of the LibreOffice Instance.

        |lo_unsafe|

        Returns:
            str: Language string such as 'en-US'
        """

        try:
            # this value is not expected to change in multi document mode.
            return cls._language
        except AttributeError:
            # sourcery skip: use-or-for-fallback
            lang = cls.get_config(node_str="ooLocale", node_path="/org.openoffice.Setup/L10N")
            if not lang:
                lang = cls.get_config(
                    node_str="ooSetupSystemLocale",
                    node_path="/org.openoffice.Setup/L10N",
                )
            if not lang:
                # default to en-us
                lang = "en-US"
            cls._language = str(lang)
        return cls._language

    @classproperty
    def language_locale(cls) -> Locale:
        """
        Gets the Current Language ``Locale`` of the LibreOffice Instance.

        |lo_unsafe|

        Returns:
            Locale: ``Locale`` object.

        .. versionadded:: 0.9.4
        """

        try:
            # this value is not expected to change in multi document mode.
            return cls._language_locale
        except AttributeError:
            cls._language_locale = cls.parse_language_code(cls.language)
        return cls._language_locale

    @classproperty
    def version(cls) -> str:
        """
        Gets the running LibreOffice version.

        |lo_unsafe|

        Returns:
            str: version as string such as ``"7.3.4.2"``

        Note:
            This property only works after Office is Loaded.
        """

        try:
            return cls._version
        except AttributeError:
            # ooSetupVersionAboutBox long "7.3.4.2"
            # ooSetupVersion short "7.3"
            lang = cls.get_config(node_str="ooSetupVersionAboutBox")
            cls._version = str(lang)
        return cls._version

    @classproperty
    def version_info(cls) -> Tuple[int, ...]:
        """
        Gets the running LibreOffice version.

        |lo_unsafe|

        Returns:
            tuple: version as tuple such as ``(7, 3, 4, 2)``

        Note:
            This property only works after Office is Loaded.
        """

        try:
            return cls._version_info
        except AttributeError:
            cls._version_info = tuple(int(s) for s in cls.version.split("."))
        return cls._version_info


def _del_cache_attrs(source: object, e: EventArgs) -> None:
    # clears Write Attributes that are dynamically created
    data_attrs = ("_language", "_language_locale", "_version", "_version_info")
    for attr in data_attrs:
        if hasattr(Info, attr):
            delattr(Info, attr)


# subscribe to events that warrant clearing cached attribs
_Events().on(LoNamedEvent.RESET, _del_cache_attrs)

__all__ = ("Info",)
