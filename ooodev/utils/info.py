# coding: utf-8
# Python conversion of Info.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
from __future__ import annotations
from enum import IntFlag
from datetime import datetime
from pathlib import Path
import mimetypes
from typing import TYPE_CHECKING, Tuple, List, cast, overload, Optional
import uno
from .sys_info import SysInfo

from com.sun.star.awt import XToolkit
from com.sun.star.beans import XHierarchicalPropertySet
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XContentEnumerationAccess
from com.sun.star.container import XNameAccess
from com.sun.star.container import XNameContainer
from com.sun.star.document import XDocumentPropertiesSupplier
from com.sun.star.document import XTypeDetection
from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.lang import XServiceInfo
from com.sun.star.reflection import XIdlReflection
from com.sun.star.style import XStyleFamiliesSupplier
from com.sun.star.util import XChangesBatch

if TYPE_CHECKING:
    from com.sun.star.awt import FontDescriptor
    from com.sun.star.beans import PropertyValue
    from com.sun.star.beans import XPropertyContainer
    from com.sun.star.deployment import XPackageInformationProvider
    from com.sun.star.document import XDocumentProperties
    from com.sun.star.lang import XTypeProvider
    from com.sun.star.reflection import XIdlMethod

from . import lo as mLo
from . import file_io as mFileIO
from . import props as mProps
from . import date_time_util as mDate
from ..exceptions import ex as mEx
from .type_var import PathOrStr


class Info:

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
        """This filter will not be shown in the dialog box for chosing a filter in case Apache OpenOffice was not able to detect one"""
        READONLY = 0x00010000
        """All documents imported by this filter will automatically be in read-only state"""
        PREFERRED = 0x10000000
        """The filter is preferred in case of multiple filters for the same file type exist in the configuration"""
        THIRDPARTYFILTER = 0x00080000
        """
        The filter is a UNO component filter, as opposed to the internal C++ filters.
        This is an artefact that will vanish over time.

        AKA: 3RDPARTYFILTER
        """

    @staticmethod
    def get_fonts() -> Tuple[FontDescriptor, ...] | None:
        """
        Gets fonts

        Returns:
            Tuple[FontDescriptor, ...] | None: Fonts if found; Othwrwise, None
        """
        xtoolkit = mLo.Lo.create_instance_mcf(XToolkit, "com.sun.star.awt.Toolkit")
        device = xtoolkit.createScreenCompatibleDevice(0, 0)
        if device is None:
            print("Could not access graphical output device")
            return None
        return device.getFontDescriptors()

    @classmethod
    def get_font_names(cls) -> List[str] | None:
        """
        Gets font names

        Returns:
            List[str] | None: Font names if found; Othwrwise, None
        """
        fds = cls.get_fonts()
        if fds is None:
            return None

        names_set = set()
        for name in fds:
            names_set.add(name)
        names = list(names_set)
        names.sort()
        return names

    @staticmethod
    def get_font_mono_name() -> str:
        """
        Gets a general font such as ``Courier New`` (windows) or ``Liberation Mono``

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
        Gets a general font such as ``Times New Roman`` (windows) or ``Liberation Serif``

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
        Get registered modifications path

        Returns:
            str: registered modificatoins path
        """
        user_cfg_dir = mFileIO.FileIO.url_to_path(cls.get_paths("UserConfig"))
        parent_path = Path(user_cfg_dir).parent
        return str(parent_path / cls.REG_MOD_FNM)

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str) -> str:
        """
        Gets value from 'registrymodifications.xcu'

        Args:
            item (str): item name
            prop (str): property value

        Raises:
            ValueError: if unable to get value

        Returns:
            str: value from 'registrymodifications.xcu'. e.g. "Writer/MailMergeWizard" null, "MailAddress"
        """
        ...

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str, node: str) -> str:
        """
        Gets value from 'registrymodifications.xcu'

        Args:
            item (str): item name
            prop (str): property value
            node (str): node

        Raises:
            ValueError: if unable to get value

        Returns:
            str: value from 'registrymodifications.xcu'. e.g. "Writer/MailMergeWizard" null, "MailAddress"
        """
        ...

    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str, node: Optional[str] = None) -> str:
        """
        Gets value from 'registrymodifications.xcu'

        Args:
            item (str): item name
            prop (str): property value
            node (str): node

        Raises:
            ValueError: if unable to get value

        Returns:
            str: value from 'registrymodifications.xcu'. e.g. "Writer/MailMergeWizard" null, "MailAddress"
        """
        # return value from "registrymodifications.xcu"
        # e.g. "Writer/MailMergeWizard" null, "MailAddress"
        # e.g. "Logging/Settings", "org.openoffice.logging.sdbc.DriverManager", "LogLevel"
        #
        # This xpath doesn't deal with all cases in the XCU file, which sometimes
        # has many node levels between the item and the prop
        if mLo.Lo.is_macro_mode:
            raise mEx.NotSupportedMacroModeError("get_reg_item_prop() is not supported from a macro")
        try:
            from lxml import etree as XML_ETREE
        except ImportError as e:
            raise Exception("get_reg_item_prop() requires lxml python package") from e

        try:
            _xml_parser = XML_ETREE.XMLParser(remove_blank_text=True)
            fnm = cls.get_reg_mods_path()
            tree: XML_ETREE._ElementTree = XML_ETREE.parse(fnm, parser=_xml_parser)

            if node is None:
                xpath = f"//item[@oor:path='/org.openoffice.Office.{item}']/prop[@oor:name='{prop}']"
            else:
                xpath = f"']/prop[@oor:name='{item}']/node[@oor:name='{node}']/prop[@oor:name='{prop}']"
            value = tree.xpath(xpath)
            if value is None:
                raise Exception("Item Property not found")
            else:
                value = str(value).strip()
                if value == "":
                    raise Exception("Item Property is white space (?)")
            return value
        except Exception as e:
            raise ValueError("unable to get value from registrymodifications.xcu") from e

    @overload
    @classmethod
    def get_config(cls, node_str: str) -> object:
        """
        Get config

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
    def get_config(cls, node_str: str, node_path: str) -> object:
        """
        Get config

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
    def get_config(cls, node_str: str, node_path: Optional[str] = None) -> object:
        """
        Get config

        Args:
            node_str (str): node string
            node_path (str): node_path

        Raises:
            ConfigError: if unable to get config

        Returns:
            object: config
        """
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
        props = cls.get_config_props(node_path)
        return mProps.Props.get_property(xprops=props, name=node_str)

    @classmethod
    def _get_config2(cls, node_str: str) -> object:
        for node_path in cls.NODE_PATHS:
            return cls._get_config1(node_str=node_str, node_path=node_path)
        raise ValueError(f"{node_str} not found")

    @staticmethod
    def get_config_props(node_path: str) -> XPropertySet:
        """
        Get config properties

        Args:
            node_path (str): nod path

        Raises:
            PropertyError: if unable to get get property set

        Returns:
            XPropertySet: Property set
        """
        try:
            con_prov = mLo.Lo.create_instance_mcf(
                XMultiServiceFactory, "com.sun.star.configuration.ConfigurationProvider"
            )
            if con_prov is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            p = mProps.Props.make_props(nodepath=node_path)
            ca = con_prov.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", p)
            ps = mLo.Lo.qi(XPropertySet, ca)
            if ps is None:
                raise mEx.MissingInterfaceError(XPropertySet)
            return ps
        except Exception as e:
            raise mEx.PropertyError(node_path, f"Unable to access config properties for\n\n  '{node_path}'") from e

    @staticmethod
    def get_paths(setting: str) -> str:
        """
        Gets access to LO's prefefined paths.

        Args:
            setting (str): property value

        Raises:
            ValueError: if unable to get paths

        Returns:
            str: paths

        Note:
            There are two different groups of properties.
            One group stores only a single path and the other group stores two or
            more paths - separated by a semicolon.

            Some setting values (as listed in the OpenOffice docs for PathSettings)

                - Addin
                - AutoCorrect
                - AutoText
                - Backup
                - Basic
                - Bitmap
                - Config
                - Dictionary
                - Favorite
                - Filter
                - Gallery
                - Graphic
                - Help
                - Linguistic
                - Module
                - Palette
                - Plugin
                - Storage
                - Temp
                - Template
                - UIConfig
                - UserConfig
                - UserDictionary (deprecated)
                - Work

        See Also:
            :py:meth:`Info.get_dirs`

            `Wiki Path Settings <https://wiki.openoffice.org/w/index.php?title=Documentation/DevGuide/OfficeDev/Path_Settings>`_
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
            prop_set = mLo.Lo.create_instance_mcf(XPropertySet, "com.sun.star.util.PathSettings")
            if prop_set is None:
                raise mEx.MissingInterfaceError(XPropertySet)
            result = prop_set.getPropertyValue(setting)
            if result is None:
                raise ValueError(f"getPropertyValue() for {setting} yielded None")
            return str(result)
        except Exception as e:
            raise ValueError(f"Could not find paths for: {setting}") from e

    @classmethod
    def get_dirs(cls, setting: str) -> List[str] | None:
        """
        Gets dirs paths from settings

        Args:
            setting (str): setting

        Returns:
            List[str] | None: List of paths if found; Otherwise, None

        See Also:
            :py:meth:`~Info.get_paths`

            `Wiki Path Settings <https://wiki.openoffice.org/w/index.php?title=Documentation/DevGuide/OfficeDev/Path_Settings>`_
        """
        try:
            paths = cls.get_paths(setting)
        except ValueError:
            print(f"Cound not find paths for '{setting}'")
            return None
        paths_arr = paths.split(";")
        if len(paths_arr) == 0:
            print(f"Cound not split paths for '{setting}'")
            return [mFileIO.FileIO.uri_to_path(paths)]
        dirs = []
        for el in paths_arr:
            dirs.append(mFileIO.FileIO.uri_to_path(el))
        return dirs

    @classmethod
    def get_office_dir(cls) -> str:
        """
        Gets file path to the office dir.
        e.g. ``"C:\Program Files (x86)\LibreOffice 7"``

        Raises:
            ValueError: if unable to obtain office path.

        Returns:
            str: Path as string
        """
        try:
            addin_dir = cls.get_paths("Addin")

            addin_path = mFileIO.FileIO.uri_to_path(addin_dir)
            #   e.g.  C:\Program Files (x86)\LibreOffice 6\program\addin
            try:
                idx = addin_path.index("program")
            except ValueError:
                print("Cound not extract office path")
                return addin_path

            p = Path(addin_path[:idx])
            return str(p)
        except Exception as e:
            raise ValueError("Unable to get office dir") from e

    @classmethod
    def get_gallery_dir(cls) -> str:
        """
        Get the first directory that contain the Gallery database and multimedia files.

        Raises:
            ValueError if unable to obtain gallery dir.

        Returns:
            str: Gallery Dir
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
        Create Configuration View

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
        Get config properties

        Args:
            node_path (str): Node path of properties

        Raises:
            mEx.ConfigError: if Unble to get config properties

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
        Sets config

        Args:
            node_path (str): node path
            node_str (str): node name
            val (object): node value

        Returns:
            bool: True on success; Otherwise, False
        """
        try:
            props = cls.set_config_props(node_path=node_path)
            if props is None:
                return False
            mProps.Props.set_property(prop_set=props, name=node_str, value=val)
            secure_change = mLo.Lo.qi(XChangesBatch, props)
            if secure_change is None:
                raise mEx.MissingInterfaceError(XChangesBatch)
            secure_change.commitChanges()
            return True
        except Exception as e:
            pass
        return False

    # =================== getting info about a document ====================

    @staticmethod
    def get_name(fnm: PathOrStr) -> str:
        """
        Gets the file's name from the supplied string minus the extension

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
        if p.suffix == "":
            print(f"No extension found for '{fnm}'")
            return p.stem
        return p.stem

    @staticmethod
    def get_ext(fnm: PathOrStr) -> str | None:
        """
        Gets file extenson without the ``.``

        Args:
            fnm (PathOrStr): file path

        Raises:
            ValueError: If fnm is empty string

        Returns:
            str | None: Extension if Found; Otherwise, None
        """
        if fnm == "":
            raise ValueError("Empty string")
        p = Path(fnm)
        # if not p.is_file():
        #     print(f"Not a file: {fnm}")
        #     return None
        if p.suffix == "":
            print(f"No extension found for '{fnm}'")
            return None
        return p.suffix[1:]

    @staticmethod
    def get_unique_fnm(fnm: PathOrStr) -> str:
        """
        If a file called fnm already exists, then a number
        is added to the name so the filename is unique

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
        Gets doc type from file path

        Args:
            fnm (PathOrStr): File Path

        Raises:
            ValueError: if Unable to get doc type

        Returns:
            str: Doc Type.
        """
        try:
            xdetect = mLo.Lo.create_instance_mcf(XTypeDetection, "com.sun.star.document.TypeDetection")
            if xdetect is None:
                raise mEx.MissingInterfaceError(XTypeDetection)
            if not mFileIO.FileIO.is_openable(fnm):
                raise mEx.UnOpenableError(fnm)
            url_str = mFileIO.FileIO.fnm_to_url(fnm)
            media_desc = [[mProps.Props.make_prop_value(name="URL", value=url_str)]]
            result = xdetect.queryTypeByDescriptor(media_desc, True)
            if result is None:
                raise mEx.UnKnownError("queryTypeByDescriptor() is an unknow result")
            return result
        except Exception as e:
            raise ValueError(f"unable to get doc type for ''{fnm}") from e

    @classmethod
    def report_doc_type(cls, doc: object) -> mLo.Lo.DocType:
        """
        Prints doc type to console and return doc type

        Args:
            doc (object): office document

        Returns:
            Lo.DocType: Document type.
        """
        doc_type = mLo.Lo.DocType.UNKNOWN
        if cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.WRITER):
            print("A Writer document")
            doc_type = mLo.Lo.DocType.WRITER
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.IMPRESS):
            print("A Impress document")
            doc_type = mLo.Lo.DocType.IMPRESS
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.DRAW):
            print("A Draw document")
            doc_type = mLo.Lo.DocType.DRAW
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.CALC):
            print("A Calc document")
            doc_type = mLo.Lo.DocType.CALC
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.BASE):
            print("A Base document")
            doc_type = mLo.Lo.DocType.BASE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.MATH):
            print("A Math document")
            doc_type = mLo.Lo.DocType.MATH
        else:
            print("Unknown document")
        return doc_type

    @classmethod
    def doc_type_service(cls, doc: object) -> mLo.Lo.Service:
        """
        Prints service type to console and return service type

        Args:
            doc (object): office document

        Returns:
            Lo.Service: Serive type
        """
        if cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.WRITER):
            print("A Writer document")
            return mLo.Lo.Service.WRITER
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.IMPRESS):
            print("A Impress document")
            return mLo.Lo.Service.IMPRESS
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.DRAW):
            print("A Draw document")
            return mLo.Lo.Service.DRAW
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.CALC):
            print("A Calc document")
            return mLo.Lo.Service.CALC
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.BASE):
            print("A Base document")
            return mLo.Lo.Service.BASE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.Service.MATH):
            print("A Math document")
            return mLo.Lo.Service.MATH
        else:
            print("Unknown document")
            return mLo.Lo.Service.UNKNOWN

    @staticmethod
    def is_doc_type(obj: object, doc_type: mLo.Lo.Service) -> bool:
        """
        Gets if doc is a particular doc type.

        Args:
            obj (object): office document
            doc_type (mLo.Lo.Service): doc type

        Returns:
            bool: True if obj matches; Otherwise, False
        """
        try:
            si = mLo.Lo.qi(XServiceInfo, obj)
            if si is None:
                return False
            return si.supportsService(str(doc_type))
        except Exception:
            return False

    @staticmethod
    def get_implementation_name(obj: object) -> str:
        """
        Gets implementaton name such as ``com.sun.star.comp.deployment.PackageInformationProvider``

        Args:
            obj (object): uno object that implements XServiceInfo

        Raises:
            ValueError: if unable to get implementation name

        Returns:
            str: implementation name
        """
        try:
            si = mLo.Lo.qi(XServiceInfo, obj)
            if si is None:
                raise mEx.MissingInterfaceError(XServiceInfo)
            return si.getImplementationName()
        except Exception as e:
            raise ValueError("Could not get service information") from e

    @staticmethod
    def get_mime_type(fnm: PathOrStr) -> str:
        """
        Get mime type for a file path

        Args:
            fnm (PathOrStr): file path

        Returns:
            str: Mime type of file if found. Defaults to 'application/octet-stream'
        """
        default = "application/octet-stream"
        mt = mimetypes.guess_type(fnm)
        if mt is None:
            print("unable to find mimeypte")
            return default
        if mt[0] is None:
            print("unable to find mimeypte")
            return default
        return str(mt[0])

    @staticmethod
    def mime_doc_type(mime_type: str) -> mLo.Lo.DocType:
        """
        Gets document type from mime type

        Args:
            mime_type (str): mime type

        Returns:
            Lo.DocType: Document type. If mime_type is unknown then 'DocType.UNKNOWN'
        """
        if mime_type is None or mime_type == "":
            return mLo.Lo.DocType.UNKNOWN
        if mime_type.find("vnd.oasis.opendocument.text") >= 0:
            return mLo.Lo.DocType.WRITER
        if mime_type.find("vnd.oasis.opendocument.base") >= 0:
            return mLo.Lo.DocType.BASE
        if mime_type.find("vnd.oasis.opendocument.spreadsheet") >= 0:
            return mLo.Lo.DocType.CALC
        if (
            mime_type.find("vnd.oasis.opendocument.graphics") >= 0
            or mime_type.find("vnd.oasis.opendocument.image") >= 0
            or mime_type.find("vnd.oasis.opendocument.chart") >= 0
        ):
            return mLo.Lo.DocType.DRAW
        if mime_type.find("vnd.oasis.opendocument.presentation") >= 0:
            return mLo.Lo.DocType.IMPRESS
        if mime_type.find("vnd.oasis.opendocument.formula") >= 0:
            return mLo.Lo.DocType.MATH
        return mLo.Lo.DocType.UNKNOWN

    @staticmethod
    def is_image_mime(mime_type: str) -> bool:
        """
        Gets if mime type is a known image type

        Args:
            mime_type (str): mime type e.g. 'application/x-openoffice-bitmap'

        Returns:
            bool: True if known mime type; Otherwise False
        """
        if mime_type is None or mime_type == "":
            return False
        if mime_type.startswith("image/"):
            return True
        if mime_type.startswith("application/x-openoffice-bitmap"):
            return True
        return False

    # ------------------------ services, interfaces, methods info ----------------------
    @overload
    @classmethod
    def get_service_names(cls) -> List[str] | None:
        """
        Gets service names

        Returns:
            List[str] | None: Service names on success; Othwriwse, None
        """
        ...

    @overload
    @classmethod
    def get_service_names(cls, service_name: str) -> List[str] | None:
        """
        Gets service names

        Args:
            service_name (str): service name

        Returns:
            List[str] | None: Service names on success; Othwriwse, None
        """
        ...

    @classmethod
    def get_service_names(cls, service_name: Optional[str] = None) -> List[str] | None:
        """
        Gets service names

        Args:
            service_name (str): service name

        Returns:
            List[str] | None: Service names on success; Othwriwse, None
        """
        if service_name is None:
            return cls._get_service_names1()
        return cls._get_service_names2(service_name=service_name)

    @staticmethod
    def _get_service_names1() -> List[str] | None:
        mc_factory = mLo.Lo.get_component_factory()
        if mc_factory is None:
            return None
        service_names = list(mc_factory.getAvailableServiceNames())
        service_names.sort()
        return service_names

    @staticmethod
    def _get_service_names2(service_name: str) -> List[str] | None:
        names: List[str] = []
        try:
            enum_access = mLo.Lo.qi(XContentEnumerationAccess, mLo.Lo.get_component_factory())
            x_enum = enum_access.createContentEnumeration(service_name)
            while x_enum.hasMoreElements():
                si = mLo.Lo.qi(XServiceInfo, x_enum.nextElement())
                names.append(si.getImplementationName())
        except Exception:
            print(f"Could not collect service names for: {service_name}")
            return None
        if len(names) == 0:
            print(f"No service names found for: {service_name}")
            return None

        names.sort()
        return names

    @staticmethod
    def get_services(obj: object) -> List[str] | None:
        """
        Gets service names

        Args:
            obj (object): obj that implements XServiceInfo

        Returns:
            List[str] | None: service names on success; Othwrwise, None
        """
        try:
            si = mLo.Lo.qi(XServiceInfo, obj)
            if si is None:
                raise mEx.MissingInterfaceError(XServiceInfo)
            names = si.getSupportedServiceNames()
            service_names = list(names)
            service_names.sort()
            return service_names
        except Exception as e:
            print("Unable to get services")
            print(f"    {e}")
        return None

    @classmethod
    def show_services(cls, obj_name: str, obj: object) -> None:
        """
        Prints services to console

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
    def support_service(obj: object, service: str) -> bool:
        """
        Gets if ``obj`` supports service

        Args:
            obj (object): Object to check for supported service
            service (string): Any UNO such as 'com.sun.star.configuration.GroupAccess'

        Returns:
            bool: True if obj supports service; Otherwise; False
        """

        if isinstance(service, str):
            srv = service
        else:
            raise TypeError(f"service is expected to be a string")
        try:
            si = mLo.Lo.qi(XServiceInfo, obj)
            if si is None:
                raise mEx.MissingInterfaceError(XServiceInfo)
            return si.supportsService(srv)
        except Exception as e:
            print("Errors ocurred in support_service(). Returning False")
            print(f"    {e}")
            pass
        return False

    @staticmethod
    def get_available_services(obj: object) -> List[str] | None:
        """
        Gets available services for obj

        Args:
            obj (object): obj that implements XMultiServiceFactory interface

        Returns:
            List[str] | None: List of services on success; Otherwise, None
        """
        services: List[str] = []
        try:
            sf = mLo.Lo.qi(XMultiServiceFactory, obj)
            if sf is None:
                raise mEx.MissingInterfaceError(XMultiServiceFactory)
            service_names = sf.getAvailableServiceNames()
            services.extend(service_names)
            services.sort()
        except Exception as e:
            print(e)
        if len(services) == 0:
            return None
        return services

    @staticmethod
    def get_interface_types(target: object) -> Tuple[object, ...] | None:
        """
        Get interface types

        Args:
            target (object): object that implements XTypeProvider interface

        Returns:
            Tuple[object, ...] | None: Tuple of interfaces on success; Othwrwise, None
        """
        try:
            tp = mLo.Lo.qi(XTypeProvider, target)
            if tp is None:
                raise mEx.MissingInterfaceError(XTypeProvider)
            types = tp.getTypes()
            return types
        except Exception as e:
            print("Unable to get interface types")
            print(f"    {e}")
        return None

    @overload
    @classmethod
    def get_interfaces(cls, target: object) -> List[str] | None:
        """
        Gets interfaces

        Args:
            target: (object): object that implements XTypeProvider

        Returns:
            List[str] | None: List of interfaces on success; Otherwise, None
        """
        ...

    @overload
    @classmethod
    def get_interfaces(cls, type_provider: XTypeProvider) -> List[str] | None:
        """
        Gets interfaces

        Args:
            type_provider (XTypeProvider): type provider

        Returns:
            List[str] | None: List of interfaces on success; Otherwise, None
        """
        ...

    @classmethod
    def get_interfaces(cls, *args, **kwargs) -> List[str] | None:
        """
        Gets interfaces

        Args:
            target: (object): object that implements XTypeProvider
            type_provider (XTypeProvider): type provider

        Returns:
            List[str] | None: List of interfaces on success; Otherwise, None
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("target", "typeProvider")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_interfaces() got an unexpected keyword argument")
            keys = ("target", "typeProvider")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("get_interfaces() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        try:
            if cls.is_type_interface(kargs[0], XTypeProvider.__pyunointerface__):
                type_provider = cast(XTypeProvider, kargs[0])
            else:
                type_provider = mLo.Lo.qi(XTypeProvider, kargs[0])
                if type_provider is None:
                    raise mEx.MissingInterfaceError(XTypeProvider)

            types = type_provider.getTypes()
            # use a set to exclude duplicate names
            names_set = set()
            for t in types:
                names_set.add(t.getTypeName())
            type_names = list(names_set)
            type_names.sort()
            return type_names
        except Exception as e:
            print("Unable to get interfaces")
            print(f"    {e}")
        return None

    @classmethod
    def show_interfaces(cls, obj_name: str, obj: object) -> None:
        """
        prints interfaces in obj to console

        Args:
            obj_name (str): Name of object for printing
            obj (object): obj that contains interfaces.
        """
        intfs = cls.get_interfaces(obj)
        if intfs is None:
            print(f"No interfaces found for {obj_name}")
            return
        print(f"{obj_name} Interfaces ({len(intfs)})")
        for s in intfs:
            print(f"  {s}")

    @staticmethod
    def get_methods(interface_name: str) -> List[str] | None:
        """
        Get Interface Methods

        Args:
            interface_name (str): name of interface

        Returns:
            List[str] | None: List of methods on success; Otherwise, None
        """
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

        reflection = mLo.Lo.create_instance_mcf(XIdlReflection, "com.sun.star.reflection.CoreReflection")
        fname = reflection.forName(interface_name)

        if fname is None:
            print(f"Could not find the interface name: {interface_name}")
            return None
        try:
            methods: Tuple[XIdlMethod, ...] = fname.getMethods()
            lst = []
            for meth in methods:
                lst.append(meth.getName())
            lst.sort()
            return lst
        except Exception as e:
            print(f"Could not get Methods for: {interface_name}")
            print(f"    {e}")
        return None

    @classmethod
    def show_methods(cls, interfce_name: str) -> None:
        """
        Prints methods to console for an interface

        Args:
            interfce_name (str): name of interface
        """
        methods = cls.get_methods(interface_name=interfce_name)
        if methods is None:
            return
        print(f"{interfce_name} Methods{len(methods)}")
        for method in methods:
            print(f"  {method}")

    # -------------------------- style info --------------------------

    @staticmethod
    def get_style_family_names(doc: object) -> List[str] | None:
        """
        Gets a list of style family names

        Args:
            doc (object): office document

        Returns:
            List[str] | None: List of style names on success; Otherwise, None
        """
        try:
            xsupplier = mLo.Lo.qi(XStyleFamiliesSupplier, doc)
            if xsupplier is None:
                raise mEx.MissingInterfaceError(XStyleFamiliesSupplier)
            name_acc = xsupplier.getStyleFamilies()
            names = name_acc.getElementNames()
            lst = list(names)
            lst.sort()
            return lst
        except Exception as e:
            print("Unable to get family style names")
            print(f"    {e}")
        return None

    @staticmethod
    def get_style_container(doc: object, family_style_name: str) -> XNameContainer:
        """
        Gets style container of document for a family of styles

        Args:
            doc (object): office document
            family_style_name (str): Family style name

        Raises:
            MissingInterfaceError: if doc is missing XStyleFamiliesSupplier interface
            MissingInterfaceError: if unable to obtain XNameContainer interface

        Returns:
            XNameContainer: Style Family container
        """
        xsupplier = mLo.Lo.qi(XStyleFamiliesSupplier, doc)
        if xsupplier is None:
            raise mEx.MissingInterfaceError(XStyleFamiliesSupplier)
        name_acc = xsupplier.getStyleFamilies()
        xcontianer = mLo.Lo.qi(XNameContainer, name_acc.getByName(family_style_name))
        if xcontianer is None:
            raise mEx.MissingInterfaceError(XNameContainer)
        return xcontianer

    @classmethod
    def get_style_names(cls, doc: object, family_style_name: str) -> List[str] | None:
        """
        Gets a list of style names

        Args:
            doc (object): office document
            family_style_name (str): name of family style

        Returns:
            List[str] | None: List of style names on success; Otherwise, None
        """
        try:
            style_container = cls.get_style_container(doc=doc, family_style_name=family_style_name)
            names = style_container.getElementNames()
            lst = list(names)
            lst.sort()
            return lst
        except Exception as e:
            print("Could not access style names")
            print(f"    {e}")
        return None

    @classmethod
    def get_style_props(cls, doc: object, family_style_name: str, prop_set_nm: str) -> XPropertySet:
        """
        Get style properties for a family of styles

        Args:
            doc (object): office document
            family_style_name (str): name of family style
            prop_set_nm (str): property set name

        Raises:
            MissingInterfaceError: if a required interface cannot be obtained.

        Returns:
            XPropertySet: Property set
        """
        style_container = cls.get_style_container(doc, family_style_name)
        #       container is a collection of named property sets
        name_props = mLo.Lo.qi(XPropertySet, style_container.getByName(prop_set_nm))
        if name_props is None:
            raise mEx.MissingInterfaceError(XPropertySet)
        return name_props

    @classmethod
    def get_page_style_props(cls, doc: object) -> XPropertySet:
        """
        Gets style properties for page styles

        Args:
            doc (object): office docs

        Raises:
            MissingInterfaceError: if a required interface cannot be obtained.

        Returns:
            XPropertySet: property set
        """
        return cls.get_style_props(doc, "PageStyles", "Standard")

    @classmethod
    def get_paragraph_style_props(cls, doc: object) -> XPropertySet:
        """
        Gets style properties for paragraph styles

        Args:
            doc (object): office docs

        Raises:
            MissingInterfaceError: if a required interface cannot be obtained.

        Returns:
            XPropertySet: property set
        """
        return cls.get_style_props(doc, "ParagraphStyles", "Standard")

    # ----------------------------- document properties ----------------------

    @staticmethod
    def str_date_time(dt: datetime) -> str:
        """
        returns a formated date and time as string

        Args:
            dt (datetime): date time

        Returns:
            str: formatted date string such as 'Jun 05, 2022 20:15'
        """
        return mDate.DateUtil.date_time_str(dt=dt)

    @classmethod
    def print_doc_properties(cls, doc: object) -> None:
        """
        Prints document properties to console

        Args:
            doc (object): office document
        """
        try:
            doc_props_supp = mLo.Lo.qi(XDocumentPropertiesSupplier, doc)
            if doc_props_supp is None:
                raise mEx.MissingInterfaceError(XDocumentPropertiesSupplier)
            dps = doc_props_supp.getDocumentProperties()
            cls.print_doc_props(dps=dps)
            ud_props = dps.getUserDefinedProperties()
            mProps.Props.show_obj_props("UserDefined Info", ud_props)
        except Exception as e:
            print("Unable to get doc properties")
            print(f"    {e}")
        return

    @classmethod
    def print_doc_props(cls, dps: XDocumentProperties) -> None:
        """
        Prints doc properties to console

        Args:
            dps (XDocumentProperties): document properties.

        See Also:
            :py:meth:`~Info.print_doc_properties`
        """
        print("Document Properties Info")
        print("  Author: " + dps.Author)
        print("  Title: " + dps.Title)
        print("  Subject: " + dps.Subject)
        print("  Description: " + dps.Description)
        print("  Generator: " + dps.Generator)

        keys: List[str] = dps.getKeywords()
        print("  Keywords: ")
        for keyword in keys:
            print(f"  {keyword}")

        print("  Modified by: " + dps.ModifiedBy)
        print("  Printed by: " + dps.PrintedBy)
        print("  Template Name: " + dps.TemplateName)
        print("  Template URL: " + dps.TemplateURL)
        print("  Autoload URL: " + dps.AutoloadURL)
        print("  Default Target: " + dps.DefaultTarget)

        l = dps.Language
        print(f"  Locale: {l.Language}; {l.Country}; {l.Variant}")

        print("  Modification Date: " + cls.str_date_time(dps.ModificationDate))
        print("  Creation Date: " + cls.str_date_time(dps.CreationDate))
        print("  Print Date: " + cls.str_date_time(dps.PrintDate))
        print("  Template Date: " + cls.str_date_time(dps.TemplateDate))

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
    def set_doc_props(doc: object, subject: str, title: str, author: str) -> None:
        """
        Set document properties for subject, title, author

        Args:
            doc (object): office document
            subject (str): subject
            title (str): title
            author (str): author

        Raises:
            PropertiesError: If unable to set properties
        """
        try:
            dp_supplier = mLo.Lo.qi(XDocumentPropertiesSupplier, doc)
            if dp_supplier is None:
                raise mEx.MissingInterfaceError(XDocumentPropertiesSupplier)
            doc_props = dp_supplier.getDocumentProperties()
            doc_props.Subject = subject
            doc_props.Title = title
            doc_props.Author = author
        except Exception as e:
            raise mEx.PropertiesError("Unable to set doc properties") from e

    @staticmethod
    def get_user_defined_props(doc: object) -> XPropertyContainer:
        """
        Gets user defined properties

        Args:
            doc (object): office document

        Raises:
            PropertiesError: if unable to access properties

        Returns:
            XPropertyContainer: Property container
        """
        try:
            dp_supplier = mLo.Lo.qi(XDocumentPropertiesSupplier, doc)
            if dp_supplier is None:
                raise mEx.MissingInterfaceError(XDocumentPropertiesSupplier)
            dps = dp_supplier.getDocumentProperties()
            return dps.getUserDefinedProperties()
        except Exception as e:
            raise mEx.PropertiesError("Unable to get user defined props") from e

    # ----------- installed package info -----------------

    @staticmethod
    def get_pip() -> XPackageInformationProvider:
        """
        Gets Package Information Provider

        Raises:
            MissingInterfaceError: if unable to obtain XPackageInformationProvider interface

        Returns:
            XPackageInformationProvider: Package Information Provider
        """
        ctx = mLo.Lo.get_context()
        pip = mLo.Lo.qi(
            XPackageInformationProvider,
            ctx.getValueByName("/singletons/com.sun.star.deployment.PackageInformationProvider"),
        )
        if pip is None:
            raise mEx.MissingInterfaceError(XPackageInformationProvider)
        return pip
        # return pip.get(mLo.Lo.get_context())

    @classmethod
    def list_extensions(cls) -> None:
        """
        Prints extensions to console
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
    def get_extension_info(cls, id: str) -> Tuple[str, ...] | None:
        """
        Gets infor for an installed extension in LibreOffice.

        Args:
            id (str): Extension id

        Returns:
            Tuple[str, ...] | None: Extension info on success; Otherwise, None
        """
        try:
            pip = cls.get_pip()
        except mEx.MissingInterfaceError:
            print("No package info provider found")
            return
        exts_tbl = pip.getExtensionList()
        mLo.Lo.print_table("Extension", exts_tbl)
        for el in exts_tbl:
            if el[0] == id:
                return el

        print(f"Extension {id} is not found")
        return None

    @classmethod
    def get_extension_loc(cls, id: str) -> str | None:
        """
        Gets location for an installed extension in LibreOffice

        Args:
            id (str): Extension id

        Returns:
            str | None: Extension location on success; Otherwise, None
        """
        try:
            pip = cls.get_pip()
        except mEx.MissingInterfaceError:
            print("No package info provider found")
            return None
        return pip.getPackageLocation(id)

    @staticmethod
    def get_filter_names() -> Tuple[str, ...] | None:
        """
        Gets filter names

        Returns:
            Tuple[str, ...] | None: Filter names on success; Otherwise, None
        """
        na = mLo.Lo.create_instance_mcf(XNameAccess, "com.sun.star.document.FilterFactory")
        if na is None:
            print("No Filter factory found")
            return None
        return na.getElementNames()

    @staticmethod
    def get_filter_props(filter_nm: str) -> List[PropertyValue] | None:
        """
        Gets filter properties

        Args:
            filter_nm (str): Filter Name

        Returns:
            List[PropertyValue] | None: List of Properties on success; Otherwise, None
        """
        na = mLo.Lo.create_instance_mcf(XNameAccess, "com.sun.star.document.FilterFactory")
        if na is None:
            print("No Filter factory found")
            return None
        result = na.getByName(filter_nm)
        if result is None:
            print(f"No props for filter: {filter_nm}")
            return None
        return list(result)

    @classmethod
    def is_import(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.IMPORT`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.IMPORT) == cls.Filter.IMPORT

    @classmethod
    def is_export(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.EXPORT`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.EXPORT) == cls.Filter.EXPORT

    @classmethod
    def is_template(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.TEMPLATE`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.TEMPLATE) == cls.Filter.TEMPLATE

    @classmethod
    def is_internal(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.INTERNAL`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.INTERNAL) == cls.Filter.INTERNAL

    @classmethod
    def is_template_path(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.TEMPLATEPATH`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.TEMPLATEPATH) == cls.Filter.TEMPLATEPATH

    @classmethod
    def is_own(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.OWN`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.OWN) == cls.Filter.OWN

    @classmethod
    def is_alien(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.ALIEN`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.ALIEN) == cls.Filter.ALIEN

    @classmethod
    def is_default(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.DEFAULT`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.DEFAULT) == cls.Filter.DEFAULT

    @classmethod
    def is_support_selection(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.SUPPORTSSELECTION`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.SUPPORTSSELECTION) == cls.Filter.SUPPORTSSELECTION

    @classmethod
    def is_not_in_file_dialog(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.NOTINFILEDIALOG`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.NOTINFILEDIALOG) == cls.Filter.NOTINFILEDIALOG

    @classmethod
    def is_not_in_chooser(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.NOTINCHOOSER`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.NOTINCHOOSER) == cls.Filter.NOTINCHOOSER

    @classmethod
    def is_read_only(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.READONLY`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.READONLY) == cls.Filter.READONLY

    @classmethod
    def is_third_party_filter(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.THIRDPARTYFILTER`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.THIRDPARTYFILTER) == cls.Filter.THIRDPARTYFILTER

    @classmethod
    def is_preferred(cls, filter_flags: Filter) -> bool:
        """
        Gets if filter flags has ``Filter.PREFERRED`` flag set

        Args:
            filter_flags (Filter): Flags

        Returns:
            bool: True if flag is set; Otherwise, False
        """
        return (filter_flags & cls.Filter.PREFERRED) == cls.Filter.PREFERRED

    @staticmethod
    def is_type_struct(obj: object, type_name: str) -> bool:
        """
        Gets if an object is a Uno Struct of matching type.

        Args:
            obj (object): Object to test if is struct
            type_name (str): Type string such as 'com.sun.star.table.CellRangeAddress'

        Returns:
            bool: True if 'obj' is struct and 'obj' matches 'type_name'; Otherwise, False
        """
        if obj is None:
            return False
        if hasattr(obj, "typeName"):
            return obj.typeName == type_name
        return False

    @staticmethod
    def is_type_interface(obj: object, type_name: str) -> bool:
        """
        Gets if an object is a Uno interface of matching type.

        Args:
            obj (object): Object to test if is interface
            type_name (str): Type string such as 'com.sun.star.uno.XInterface'

        Returns:
            bool: True if 'obj' is interface and 'obj' matches 'type_name'; Otherwise, False
        """
        if obj is None:
            return False
        if hasattr(obj, "__pyunointerface__"):
            return obj.__pyunointerface__ == type_name
        elif hasattr(obj, "queryInterface"):
            uno_t = uno.getTypeByName(type_name)
            q_obj = obj.queryInterface(uno_t)
            if q_obj is not None:
                return True
        return False

    @staticmethod
    def is_type_enum(obj: uno.Enum, type_name: str) -> bool:
        """
        Gets if an object is a Uno enum of matching type.

        Args:
            obj (object): Object to test if is uno enum
            type_name (str): Type string such as 'com.sun.star.sheet.GeneralFunction'

        Returns:
            bool: True if 'obj' is uno enum and 'obj' matches 'type_name'; Otherwise, False
        """
        if obj is None:
            return False
        if hasattr(obj, "typeName"):
            return obj.typeName == type_name
        return False

    @classmethod
    def get_type_name(cls, obj: object) -> str | None:
        """
        Gets type name such as ``com.sun.star.table.TableSortField`` from uno object.

        Args:
            obj (object): Uno object

        Returns:
            str | None: Full type name if found; Othwrwise; None
        """
        if hasattr(obj, "typeName"):
            return obj.typeName
        if hasattr(obj, "__ooo_full_ns__"):
            # ooouno object
            return obj.__ooo_full_ns__
        if hasattr(obj, "__pyunointerface__"):
            return obj.__pyunointerface__
        return None
