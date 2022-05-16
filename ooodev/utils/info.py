# coding: utf-8
# Python conversion of Info.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
from __future__ import annotations
import sys
from datetime import datetime
from pathlib import Path
import mimetypes
from typing import TYPE_CHECKING, Tuple, List, overload, Optional
from lxml import etree as ET
import uno
from . import props as m_props
from .sys_info import SysInfo


Props = m_props.Props

if TYPE_CHECKING:
    from com.sun.star.awt import FontDescriptor
    from com.sun.star.awt import XToolkit
    from com.sun.star.beans import PropertyValue
    from com.sun.star.beans import XHierarchicalPropertySet
    from com.sun.star.beans import XPropertyContainer
    from com.sun.star.beans import XPropertySet
    from com.sun.star.container import XContentEnumerationAccess
    from com.sun.star.container import XNameContainer
    from com.sun.star.container import XNameAccess
    from com.sun.star.deployment import XPackageInformationProvider
    from com.sun.star.deployment import PackageInformationProvider
    from com.sun.star.document import XTypeDetection
    from com.sun.star.document import XDocumentPropertiesSupplier
    from com.sun.star.lang import XMultiServiceFactory
    from com.sun.star.lang import XServiceInfo
    from com.sun.star.lang import XTypeProvider
    from com.sun.star.reflection import CoreReflection
    from com.sun.star.reflection import XIdlMethod
    from com.sun.star.style import XStyleFamiliesSupplier
    from com.sun.star.uno import XInterface
    from com.sun.star.util import XChangesBatch

from . import lo as mLo
from . import file_io as mFileIO
from . import props as mProps

if sys.version_info >= (3, 10):
    from typing import Union
else:
    from typing_extensions import Union

_xml_parser = ET.XMLParser(remove_blank_text=True)


class Info:

    REG_MOD_FNM = "registrymodifications.xcu"
    NODE_PRODUCT = "/org.openoffice.Setup/Product"
    NODE_L10N = "/org.openoffice.Setup/L10N"
    NODE_OFFICE = "/org.openoffice.Setup/Office"

    NODE_PATHS = [NODE_PRODUCT, NODE_L10N]
    # MIME_FNM = "mime.types"

    # from https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Properties_of_a_Filter
    IMPORT = 0x00000001
    EXPORT = 0x00000002
    TEMPLATE = 0x00000004
    INTERNAL = 0x00000008

    TEMPLATEPATH = 0x00000010
    OWN = 0x00000020
    ALIEN = 0x00000040

    DEFAULT = 0x00000100
    SUPPORTSSELECTION = 0x00000400

    NOTINFILEDIALOG = 0x00001000
    NOTINCHOOSER = 0x00002000

    READONLY = 0x00010000
    THIRDPARTYFILTER = 0x00080000

    PREFERRED = 0x10000000

    @staticmethod
    def get_fonts() -> Tuple[FontDescriptor, ...] | None:
        xtoolkit: XToolkit = mLo.Lo.create_instance_mcf("com.sun.star.awt.Toolkit")
        device = xtoolkit.createScreenCompatibleDevice(0, 0)
        if device is None:
            print("Could not access graphical output device")
            return None
        return device.getFontDescriptors()

    @classmethod
    def get_font_name(cls) -> List[str] | None:
        fds = cls.get_fonts()
        if fds is None:
            return None

        names_set = set()
        for name in fds:
            names_set.add(name)
        names = list(names_set)
        names.sort()
        return names

    @classmethod
    def get_font_mono_name() -> str:
        """
        Gets a general font such as ``Courier New`` (windows) or ``Liberation Mono``

        Returns:
            str: Font Name

        See Also:
            `Fonts <https://wiki.documentfoundation.org/Fonts>`_ on Document Foundation’s wiki/
        """
        pf = SysInfo.get_platform()
        if pf == SysInfo.PlatformEnum.WINDOWS:
            return "Courier New"
        else:
            return "Liberation Mono"  # Metrically compatible with Courier New

    @classmethod
    def get_font_general_name() -> str:
        """
        Gets a general font such as ``Times New Roman`` (windows) or ``Liberation Serif``

        Returns:
            str: Font Name

        See Also:
            `Fonts <https://wiki.documentfoundation.org/Fonts>`_ on Document Foundation’s wiki/
        """
        pf = SysInfo.get_platform()
        if pf == SysInfo.PlatformEnum.WINDOWS:
            return "Times New Roman"
        else:
            return "Liberation Serif"  # Metrically compatible with Times New Roman

    @classmethod
    def get_reg_mods_path(cls) -> str | None:
        user_cfg_dir = mFileIO.FileIO.url_to_path(cls.get_paths("UserConfig"))

        try:
            parent_path = Path(user_cfg_dir).parent
            return str(parent_path / cls.REG_MOD_FNM)
        except Exception as e:
            print(f"Coul not parse '{user_cfg_dir}'")
            print(f"    {e}")
        return None

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str) -> str | None:
        ...

    @overload
    @classmethod
    def get_reg_item_prop(cls, item: str, prop: str, node: str) -> str | None:
        ...

    @classmethod
    def get_reg_item_prop(
        cls, item: str, prop: str, node: Optional[str] = None
    ) -> str | None:
        # return value from "registrymodifications.xcu"
        # e.g. "Writer/MailMergeWizard" null, "MailAddress"
        # e.g. "Logging/Settings", "org.openoffice.logging.sdbc.DriverManager", "LogLevel"
        #
        # This xpath doesn't deal with all cases in the XCU file, which sometimes
        # has many node levels between the item and the prop
        # Returns null if no value is found, or it's only an empty string.

        fnm = cls.get_reg_mods_path()
        if fnm is None:
            return None
        value = None

        try:
            tree: ET._ElementTree = ET.parse(fnm, parser=_xml_parser)

            if node is None:
                xpath = f"//item[@oor:path='/org.openoffice.Office.{item}']/prop[@oor:name='{prop}']"
            else:
                xpath = f"']/prop[@oor:name='{item}']/node[@oor:name='{node}']/prop[@oor:name='{prop}']"
            value = tree.xpath(xpath)
            if value is None or value == "":
                print("Item Property not founc")
                value = None
            else:
                value = str(value).strip()
                if value == "":
                    print("Item Property is white space (?)")
                    value = None
        except Exception as e:
            print(e)
        return value

    @overload
    def get_config(node_str: str) -> str | None:
        ...

    @overload
    def get_config(node_str: str, node_path: str) -> object | None:
        ...

    @classmethod
    def get_config(cls, node_str: str, node_path: Optional[str] = None):
        if node_path is None:
            return cls._get_config2(node_str=node_str)
        return cls._get_config1(node_str=node_str, node_path=node_path)

    @classmethod
    def _get_config1(cls, node_str: str, node_path: str):
        props = cls.get_config_props(node_path)
        if props is None:
            return None
        return mProps.Props.get_property(x_props=props, name=node_str)

    @classmethod
    def _get_config2(cls, node_str: str):
        for node_path in cls.NODE_PATHS:
            info = cls._get_config1(node_str=node_str, node_path=node_path)
            if info is not None:
                return info
        print(f"No configuration info for {node_str}")
        return None

    @staticmethod
    def get_config_props(node_path: str) -> XPropertySet | None:
        con_prov: XMultiServiceFactory = mLo.Lo.create_instance_mcf(
            "com.sun.star.configuration.ConfigurationProvider"
        )
        if con_prov is None:
            print("Could not create configuration provider")
            return None
        p = mProps.Props.make_props(nodepath=node_path)
        try:
            return con_prov.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess", p
            )
        except Exception as e:
            print(f"Unable to access config properties for\n\n  '{node_path}'")
        return None

    @staticmethod
    def get_paths(setting: str) -> str:
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
        prop_set: XPropertySet = mLo.Lo.create_instance_mcf(
            "com.sun.star.util.PathSettings"
        )
        if prop_set is None:
            print("Could not access office settings")
            return None
        try:
            result = prop_set.getPropertyValue(setting)
            return str(result)
        except Exception as e:
            print(f"Could not find setting for: {setting}")
        return None

    @classmethod
    def get_dirs(cls, setting: str) -> List[str] | None:
        paths = cls.get_paths(setting)
        if paths is None:
            print(f"Cound not find paths for '{setting}'")
            return None
        paths_arr = paths.split(";")
        if len(paths_arr) == 0:
            print(f"Cound not split paths for '{setting}'")
            return [paths]
        dirs = []
        for el in paths_arr:
            dirs.append(mFileIO.FileIO.uri_to_path(el))
        return dirs

    @classmethod
    def get_office_dir(cls) -> str | None:
        """
        returns the file path to the office dir

        e.g.  'C:\Program Files (x86)\LibreOffice 6'
        """
        addin_dir = cls.get_paths("Addin")
        if addin_dir is None:
            print("Cound not find settings information")
            return None
        addin_path = mFileIO.FileIO.uri_to_path(addin_dir)
        #   e.g.  C:\Program Files (x86)\LibreOffice 6\program\addin
        try:
            idx = addin_path.index("program")
        except ValueError:
            print("Cound not extract office path")
            return addin_path

        p = Path(addin_path[:idx])
        return str(p)

    @classmethod
    def get_gallery_dir(cls) -> str | None:
        gallery_dirs = cls.get_dirs("Gallery")
        if gallery_dirs is None:
            return None
        return gallery_dirs[0]

    @classmethod
    def create_configuration_view(cls, path: str) -> XHierarchicalPropertySet | None:
        con_prov: XMultiServiceFactory = mLo.Lo.create_instance_mcf(
            "com.sun.star.configuration.ConfigurationProvider"
        )
        if con_prov is None:
            print("Could not create configuration provider")
            return None
        _props = mProps.Props.make_props(nodepath=path)
        try:
            root: XInterface = con_prov.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess", _props
            )
            cls.show_services(obj_name="ConfigurationAccess", obj=root)
            return root
        except Exception:
            return None

    # =================== update configuration settings ================

    @staticmethod
    def set_config_props(node_path: str) -> XPropertySet | None:
        con_prov: XMultiServiceFactory = mLo.Lo.create_instance_mcf(
            "com.sun.star.configuration.ConfigurationProvider"
        )
        if con_prov is None:
            print("Could not create configuration provider")
            return None
        _props = mProps.Props.make_props(nodepath=node_path)
        try:
            return con_prov.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess", _props
            )
        except Exception:
            print(f"Unable to access config update properties for\n  '{node_path}'")
        return None

    @classmethod
    def set_config(cls, node_path: str, node_str: str, val: object) -> bool:
        _props: XChangesBatch = cls.set_config_props(node_path=node_path)
        if _props is None:
            return False
        mProps.Props.set_property(prop_set=_props, name=node_str, value=val)
        try:
            _props.commitChanges()
            return True
        except Exception:
            print(f"Unable to commit config update for\n  '{node_path}'")
        return False

    # =================== getting info about a document ====================

    @staticmethod
    def get_name(fnm: str) -> str:
        """extract the file's name from the supplied string minus the extension"""
        if fnm == "":
            print(f"Zero length string")
            return fnm
        p = Path(fnm)
        if not p.is_file():
            print(f"Not a file: {fnm}")
            return fnm
        if p.suffix == "":
            print(f"No extension found for '{fnm}'")
            return fnm
        return p.stem

    @staticmethod
    def get_ext(fnm: str) -> str | None:
        """return extenson without the ``.``"""
        if fnm == "":
            print(f"Zero length string")
            return None
        p = Path(fnm)
        # if not p.is_file():
        #     print(f"Not a file: {fnm}")
        #     return None
        if p.suffix == "":
            print(f"No extension found for '{fnm}'")
            return None
        return p.suffix[1:]

    @staticmethod
    def get_unique_fnm(fnm: str) -> str:
        """
        If a file called fnm already exists, then a number
        is added to the name so the filename is unique
        """
        p = Path(fnm)
        fname = p.stem
        ext = p.suffix
        i = 1
        while p.exists():
            fnm = f"{fname}{i}{ext}"
            p = p.parent / fnm
            i += 1
        return str(p)

    @staticmethod
    def get_doc_type(fnm: str) -> str | None:
        xdetect: XTypeDetection = mLo.Lo.create_instance_mcf(
            "com.sun.star.document.TypeDetection"
        )
        if xdetect is None:
            print("No type detector reference")
            return None
        if not mFileIO.FileIO.is_openable(fnm):
            return None
        url_str = mFileIO.FileIO.fnm_to_url(fnm)
        if url_str is None:
            return None
        media_desc = [[Props.make_prop_value(name="URL", value=url_str)]]
        return xdetect.queryTypeByDescriptor(media_desc, True)

    @classmethod
    def report_doc_type(cls, doc: object) -> int:
        doc_type = mLo.Lo.UNKNOWN
        if cls.is_doc_type(obj=doc, doc_type=mLo.Lo.WRITER_SERVICE):
            print("A Writer document")
            doc_type = mLo.Lo.WRITER
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.IMPRESS_SERVICE):
            print("A Impress document")
            doc_type = mLo.Lo.IMPRESS
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.DRAW_SERVICE):
            print("A Draw document")
            doc_type = mLo.Lo.DRAW
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.CALC_SERVICE):
            print("A Calc document")
            doc_type = mLo.Lo.CALC
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.BASE_SERVICE):
            print("A Base document")
            doc_type = mLo.Lo.BASE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.MATH_SERVICE):
            print("A Math document")
            doc_type = mLo.Lo.MATH
        else:
            print("Unknown document")
        return doc_type

    @classmethod
    def doc_type_string(cls, doc: object) -> str:
        if cls.is_doc_type(obj=doc, doc_type=mLo.Lo.WRITER_SERVICE):
            print("A Writer document")
            return mLo.Lo.WRITER_SERVICE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.IMPRESS_SERVICE):
            print("A Impress document")
            return mLo.Lo.IMPRESS_SERVICE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.DRAW_SERVICE):
            print("A Draw document")
            return mLo.Lo.DRAW_SERVICE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.CALC_SERVICE):
            print("A Calc document")
            return mLo.Lo.CALC_SERVICE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.BASE_SERVICE):
            print("A Base document")
            return mLo.Lo.BASE_SERVICE
        elif cls.is_doc_type(obj=doc, doc_type=mLo.Lo.MATH_SERVICE):
            print("A Math document")
            return mLo.Lo.MATH_SERVICE
        else:
            print("Unknown document")
            return mLo.Lo.UNKNOWN_SERVICE

    @staticmethod
    def is_doc_type(obj: XServiceInfo, doc_type: str) -> bool:
        try:
            return obj.supportsService(doc_type)
        except Exception:
            return False

    @staticmethod
    def get_implementation_name(obj: XServiceInfo) -> str | None:
        try:
            return obj.getImplementationName()
        except Exception as e:
            print("Could not get service information")
            print(f"    {e}")
            return None

    @staticmethod
    def get_mime_type(fnm: str) -> str:
        mt = mimetypes.guess_type(fnm)
        if mt[0] is None:
            print("unable to find mimeypte")
            return "application/octet-stream"
        return str(mt[0])

    @staticmethod
    def mime_doc_type(mime_type: str) -> int:
        if mime_type is None:
            return mLo.Lo.UNKNOWN
        if mime_type.find("vnd.oasis.opendocument.text") >= 0:
            return mLo.Lo.WRITER
        if mime_type.find("vnd.oasis.opendocument.base") >= 0:
            return mLo.Lo.BASE
        if mime_type.find("vnd.oasis.opendocument.spreadsheet") >= 0:
            return mLo.Lo.CALC
        if (
            mime_type.find("vnd.oasis.opendocument.graphics") >= 0
            or mime_type.find("vnd.oasis.opendocument.image") >= 0
            or mime_type.find("vnd.oasis.opendocument.chart") >= 0
        ):
            return mLo.Lo.DRAW
        if mime_type.find("vnd.oasis.opendocument.presentation") >= 0:
            return mLo.Lo.IMPRESS
        if mime_type.find("vnd.oasis.opendocument.formula") >= 0:
            return mLo.Lo.MATH
        return mLo.Lo.UNKNOWN

    @staticmethod
    def is_image_mime(mime_type: str) -> bool:
        if mime_type.startswith("image/"):
            return True
        if mime_type.startswith("application/x-openoffice-bitmap"):
            return True
        return False

    # ------------------------ services, interfaces, methods info ----------------------
    @overload
    @classmethod
    def get_service_names(cls) -> List[str] | None:
        ...

    @overload
    @classmethod
    def get_service_names(cls, service_name: str) -> List[str] | None:
        ...

    @classmethod
    def get_service_names(cls, service_name: Optional[str] = None) -> List[str] | None:
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
            enum_access: XContentEnumerationAccess = mLo.Lo.get_component_factory()
            x_enum = enum_access.createContentEnumeration(service_name)
            while x_enum.hasMoreElements():
                si: XServiceInfo = x_enum.nextElement()
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
    def get_services(obj: XServiceInfo) -> List[str] | None:
        try:
            names = obj.getSupportedServiceNames()
            service_names = list(names)
            service_names.sort()
            return service_names
        except AttributeError:
            print("No XServiceInfo interface found")
        except Exception as e:
            print("Unable to get services")
            print(f"    {e}")
        return None

    @classmethod
    def show_services(cls, obj_name: str, obj: object) -> None:
        services = cls.get_services()
        if services is None:
            print(f"No supported services found for {obj_name}")
            return
        print(f"{obj_name} Supported Services ({len(services)})")
        for service in services:
            print(f"'{service}'")

    @overload
    @staticmethod
    def support_service(obj: XServiceInfo, service: XInterface) -> bool:
        """
        Gets if ``obj`` supports service

        Args:
            obj (XServiceInfo): Object to check for supported service
            service (XInterface): Any UNO interface. Intervaces start with x

        Returns:
            bool: True if obj supports service; Otherwise; False
        """
        ...

    @overload
    @staticmethod
    def support_service(obj: XServiceInfo, service: str) -> bool:
        """
        Gets if ``obj`` supports service

        Args:
            obj (XServiceInfo): Object to check for supported service
            service (string): Any UNO such as 'com.sun.star.uno.XInterface'

        Returns:
            bool: True if obj supports service; Otherwise; False
        """
        ...

    @staticmethod
    def support_service(obj: XServiceInfo, service=None) -> bool:
        srv = None
        if isinstance(service, str):
            srv = service
        else:
            try:
                srv = service.__pyunointerface__
            except AttributeError:
                print("service does not have __pyunointerface__ attribute")
                return False
        try:
            return obj.supportsService(srv)
        except AttributeError:
            print("Object does not implement XServiceInfo")
        except Exception:
            pass
        return False

    @staticmethod
    def get_available_services(obj: XMultiServiceFactory) -> List[str]:
        services: List[str] = []
        try:
            service_names = obj.getAvailableServiceNames()
            services.extend(service_names)
            services.sort()
        except Exception as e:
            print(e)
        return services

    @staticmethod
    def get_interface_types(target: XTypeProvider) -> Tuple[object, ...] | None:
        try:
            types = target.getTypes()
            return types
        except AttributeError:
            print("No XTypeProvider interface found")
        except Exception as e:
            print("Unable to get types")
            print(f"    {e}")
        return None

    @staticmethod
    def get_interfaces(type_provider: XTypeProvider) -> List[str] | None:
        try:
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
        intfs = cls.get_interfaces()
        if intfs is None:
            print(f"No interfaces found for {obj_name}")
            return
        print(f"{obj_name} Interfaces ({len(intfs)})")
        for s in intfs:
            print(f"  {s}")

    @staticmethod
    def get_methods(interface_name: str) -> List[str] | None:
        """Get Interface Methods"""
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

        reflection: CoreReflection = mLo.Lo.create_instance_mcf(
            "com.sun.star.reflection.CoreReflection"
        )
        # fname = reflection.forName('com.sun.star.uno.XInterface')
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
        methods = cls.get_methods(interface_name=interfce_name)
        if methods is None:
            return
        print(f"{interfce_name} Methods{len(methods)}")
        for method in methods:
            print(f"  {method}")

    # -------------------------- style info --------------------------

    @staticmethod
    def get_style_family_names(doc: XStyleFamiliesSupplier) -> List[str] | None:
        try:
            name_acc = doc.getStyleFamilies()
            names = name_acc.getElementNames()
            lst = list(names)
            lst.sort()
            return lst
        except AttributeError:
            print("No XStyleFamiliesSupplier interface found")
        except Exception as e:
            print("Unable to get family style names")
            print(f"    {e}")
        return None

    @staticmethod
    def get_style_container(
        doc: XStyleFamiliesSupplier, family_style_name: str
    ) -> XNameContainer | None:
        try:
            name_acc = doc.getStyleFamilies()
            return name_acc.getByName(family_style_name)
        except AttributeError:
            print("No XStyleFamiliesSupplier interface found")
        except Exception as e:
            print("Unable to get style container")
            print(f"    {e}")
        return None

    @classmethod
    def get_style_names(
        cls, doc: XStyleFamiliesSupplier, family_style_name: str
    ) -> List[str] | None:
        style_container = cls.get_style_container(
            doc=doc, family_style_name=family_style_name
        )
        if style_container is None:
            return None
        names = style_container.getElementNames()
        lst = list(names)
        lst.sort()
        return lst

    @classmethod
    def get_style_props(
        cls, doc: XStyleFamiliesSupplier, family_style_name: str, prop_set_nm: str
    ) -> XPropertySet | None:
        style_container = cls.get_style_container(doc, family_style_name)
        #       container is a collection of named property sets
        if style_container is None:
            return None
        name_props = None
        try:
            name_props = style_container.getByName(prop_set_nm)
        except Exception as e:
            print(f"Could not access style: {e}")
        return name_props

    @classmethod
    def get_page_style_props(cls, doc: XStyleFamiliesSupplier) -> XPropertySet | None:
        return cls.get_style_props(doc, "PageStyles", "Standard")

    @classmethod
    def get_paragraph_style_props(
        cls, doc: XStyleFamiliesSupplier
    ) -> XPropertySet | None:
        return cls.get_style_props(doc, "ParagraphStyles", "Standard")

    # ----------------------------- document properties ----------------------

    @staticmethod
    def str_date_time(dt: datetime) -> str:
        return dt.strftime("%b %d, %Y %H:%M")

    @classmethod
    def print_doc_properties(cls, doc: XDocumentPropertiesSupplier) -> None:
        try:
            dps = doc.getDocumentProperties()
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

        except AttributeError:
            print("No XDocumentPropertiesSupplier interface found")
        except Exception as e:
            print("Unable to get doc properties")
            print(f"    {e}")
        return

    @staticmethod
    def set_doc_props(
        doc: XDocumentPropertiesSupplier, subject: str, title: str, author: str
    ) -> None:
        """Set document properties for subject, title, author"""
        try:
            doc_props = doc.getDocumentProperties()
            doc_props.Subject = subject
            doc_props.Title = title
            doc_props.Author = author
        except AttributeError:
            print("No XDocumentPropertiesSupplier interface found")
        except Exception as e:
            print("Unable to set doc properties")
            print(f"    {e}")
        return None

    @staticmethod
    def get_user_defined_props(
        doc: XDocumentPropertiesSupplier,
    ) -> XPropertyContainer | None:
        """Set document properties for subject, title, author"""
        try:
            dps = doc.getDocumentProperties()
            return dps.getUserDefinedProperties()
        except AttributeError:
            print("No XDocumentPropertiesSupplier interface found")
        except Exception as e:
            print("Unable to get user defined props")
            print(f"    {e}")
        return None

    # ----------- installed package info -----------------

    @staticmethod
    def get_pip() -> XPackageInformationProvider:
        ctx = mLo.Lo.get_context()
        pip: PackageInformationProvider = ctx.getValueByName(
            "/singletons/com.sun.star.deployment.PackageInformationProvider"
        )
        return pip
        # return pip.get(mLo.Lo.get_context())

    @classmethod
    def list_extensions(cls) -> None:
        pip = cls.get_pip()
        if pip is None:
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
    def get_extension_info(cls, id: str) -> List[str] | None:
        pip = cls.get_pip()
        if pip is None:
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
        pip = cls.get_pip()
        if pip is None:
            print("No package info provider found")
            return
        return pip.getPackageLocation(id)

    @staticmethod
    def get_filter_names() -> Tuple[str, ...] | None:
        na: XNameAccess = mLo.Lo.create_instance_mcf(
            "com.sun.star.document.FilterFactory"
        )
        if na is None:
            print("No Filter factory found")
            return None
        return na.getElementNames()

    @staticmethod
    def get_filter_props(filter_nm: str) -> List[PropertyValue] | None:
        na: XNameAccess = mLo.Lo.create_instance_mcf(
            "com.sun.star.document.FilterFactory"
        )
        if na is None:
            print("No Filter factory found")
            return None
        result = na.getByName(filter_nm)
        if result is None:
            print(f"No props for filter: {filter_nm}")
            return None
        return list(result)

    @classmethod
    def is_import(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.IMPORT) == cls.IMPORT

    @classmethod
    def is_export(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.EXPORT) == cls.EXPORT

    @classmethod
    def is_template(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.TEMPLATE) == cls.TEMPLATE

    @classmethod
    def is_internal(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.INTERNAL) == cls.INTERNAL

    @classmethod
    def is_template_path(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.TEMPLATEPATH) == cls.TEMPLATEPATH

    @classmethod
    def is_own(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.OWN) == cls.OWN

    @classmethod
    def is_alien(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.ALIEN) == cls.ALIEN

    @classmethod
    def is_default(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.DEFAULT) == cls.DEFAULT

    @classmethod
    def is_support_selection(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.SUPPORTSSELECTION) == cls.SUPPORTSSELECTION

    @classmethod
    def is_not_in_file_dialog(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.NOTINFILEDIALOG) == cls.NOTINFILEDIALOG

    @classmethod
    def is_not_in_chooser(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.NOTINCHOOSER) == cls.NOTINCHOOSER

    @classmethod
    def is_read_only(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.READONLY) == cls.READONLY

    @classmethod
    def is_third_party_filter(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.THIRDPARTYFILTER) == cls.THIRDPARTYFILTER

    @classmethod
    def is_preferred(cls, filter_flags: int) -> bool:
        return (filter_flags & cls.PREFERRED) == cls.PREFERRED
