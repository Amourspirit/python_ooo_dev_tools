# coding: utf-8
"""make/get/set properties in an array"""
# Python conversion of Props.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
# region Imports
from __future__ import annotations
from typing import Iterable, List, Optional, Sequence, Tuple, TYPE_CHECKING, cast, overload
import uno

from com.sun.star.beans import PropertyAttribute  # const
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XNameAccess
from com.sun.star.document import XTypeDetection
from com.sun.star.ui import ItemType  # const
from com.sun.star.ui import ItemStyle  # const
from com.sun.star.uno import RuntimeException
from com.sun.star.beans import PropertyVetoException
from com.sun.star.container import XIndexAccess


if TYPE_CHECKING:
    from com.sun.star.beans import Property, PropertyValue
    from com.sun.star.beans import XMultiPropertySet

    # import module and not module content to avoid circular import issue.
    # https://stackoverflow.com/questions/22187279/python-circular-importing
    #
    # using lazy loading: https://snarky.ca/lazy-importing-in-python-3-7/
from . import lo as mLo
from . import info as mInfo

from ..exceptions import ex as mEx

# endregion Imports


class Props:
    """make/get/set properties in an array"""

    # region ------------------- make properties -----------------------
    @staticmethod
    def make_prop_value(name: Optional[str] = None, value: Optional[str] = None) -> PropertyValue:
        """
        Makes a Uno Property Value and assigns name and value if present.

        Args:
            name (Optional[str], optional): Property name
            value (Optional[str], optional): Property value

        Returns:
            PropertyValue: com.sun.star.beans.PropertyValue

        See Also:
            `LibreOffice API PropertyValue <https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1beans_1_1PropertyValue.html>`_
        """
        p: PropertyValue = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        if name is not None:
            p.Name = name
        if value is not None:
            p.Value = value
        return p

    @classmethod
    def make_bar_item(
        cls, cmd: str, item_name: str
    ) -> Tuple[PropertyValue, PropertyValue, PropertyValue, PropertyValue, PropertyValue]:
        """
        Properties for a toolbar item using a name and an image

        problem: image does not appear next to text on toolbar

        Args:
            cmd (str): Value of CommandURL
            item_name (str): Label assigned to bar

        Returns:
            Tuple[PropertyValue, PropertyValue, PropertyValue, PropertyValue, PropertyValue]: Tuple of properties. (CommandURL, Label, Type, Visible, Style)
        """
        p1 = cls.make_prop_value(name="CommandURL", value=cmd)
        p2 = cls.make_prop_value(name="Label", value=item_name)
        p3 = cls.make_prop_value(name="Type", value=ItemType.DEFAULT)
        p4 = cls.make_prop_value(name="Visible", value=True)
        p5 = cls.make_prop_value(name="Style")
        p5.Value = ItemStyle.DRAW_FLAT + ItemStyle.ALIGN_LEFT + ItemStyle.AUTO_SIZE + ItemStyle.ICON + ItemStyle.TEXT

        return (p1, p2, p3, p4, p5)

    @classmethod
    def make_props(cls, **kwargs) -> Tuple[PropertyValue, ...]:
        """
        Make Properties

        Keyword Args:
            kwargs (Dict[str, Any]): Each key, value pair is assigned to a PropertyValue.

        Returns:
            Tuple[PropertyValue, ...]: Tuple of Properties

        Note:
            String properties such as ``Zoom.Value`` can be pass by constructing a dictionary
            and passing dictionary via expansion.

            Example Expansion

            .. code-block:: python

                p_dic = {
                    "Zoom.Value": 0,
                    "Zoom.ValueSet": 28703,
                    "Zoom.Type": 1
                    }
                props = Props.make_props(**p_dic)
        """
        lst = []
        for k, v in kwargs.items():
            lst.append(cls.make_prop_value(name=k, value=v))
        return tuple(lst)

    # endregion ---------------- make properties -----------------------

    # region ------------------- uno -----------------------------------
    @staticmethod
    def any(*elements: object) -> uno.Any:
        """
        Gets a uno.Any object for elements.

        The first element determines the type for the uno.Any object.

        Raises:
            ValueError: if unable to create uno.Any object.

        Returns:
            uno.Any: uno.Any Object

        Note:
            uno.Any is usually constructed in the following manor.

            ``uno.Any("[]com.sun.star.table.TableSortField", (sort_one, sort_two)``

            This method is a shortcut.

            ``Props.any(sort_one, sort_two)``
        """
        if len(elements) == 0:
            raise ValueError("No args to create unn.Any object")
        obj = elements[0]
        if isinstance(obj, uno.Type):
            type_name = obj.typeName
            return uno.Any(obj, [*elements])
        elif isinstance(obj, str):
            type_name = "string"
        elif isinstance(obj, int):
            type_name = "short"
        else:
            type_name = mInfo.Info.get_type_name(obj)
        if type_name is None:
            raise ValueError("Unable to get type name to create uno.Any object")
        return uno.Any(f"[]{type_name}", [*elements])

    # endregion ---------------- uno -----------------------------------

    # region ------------------- set properties ------------------------
    @staticmethod
    def set_prop(props: Iterable[PropertyValue], name: str, value: object) -> bool:
        """
        Sets property value for the first property that has matching name.

        Args:
            props (Iterable[PropertyValue]): Property Values
            name (str): Property name
            value (object): Property Value

        Raises:
            TypeError: if props is None.

        Returns:
            bool: True if property matching name has been updated; Otherwise, False
        """
        if props is None:
            TypeError(f"Property array is null; cannot set {name}")
        for prop in props:
            if prop.Name == name:
                prop.Value = value
                return True
        print(f"{name} not found")
        return False

    @staticmethod
    def get_prop(props: Iterable[PropertyValue], name: str) -> object | None:
        """
        Gets property value for property that matches name.

        Args:
            props (Iterable[PropertyValue]): Properties to search
            name (str): Property name to find.

        Raises:
            TypeError: if props is None.

        Returns:
            object | None: Property value if found; Otherwise; None
        """
        if props is None:
            TypeError(f"Property array is null; cannot get {name}")
        for prop in props:
            if prop.Name == name:
                return prop.Value
        print(f"{name} not found")
        return None

    # region    set_property()
    @overload
    @staticmethod
    def set_property(obj: object, name: str, value: object) -> None:
        """
        Sets the value of the property with the specified name.

        Args:
            obj (object): object that implements XPropertySet interface
            name (str): Name of property to set value of
            value (object): property value

        Raises:
            MissingInterfaceError: If required interface cannot be obtained.
            Exception: If unable to set property value
        """
        ...

    @overload
    @staticmethod
    def set_property(prop_set: XPropertySet, name: str, value: object) -> None:
        """
        Sets the value of the property with the specified name.

        Args:
            prop_set (XPropertySet): Property set
            name (str): Name of property to set value of
            value (object): property value

        Raises:
            Exception: If unable to set property value
        """
        ...

    @staticmethod
    def set_property(*args, **kwargs) -> None:
        """
        Sets the value of the property with the specified name.

        Args:
            obj (object): object that implements XPropertySet interface
            prop_set (XPropertySet): Property set
            name (str): Name of property to set value of
            value (object): property value

        Raises:
            MissingInterfaceError: If required interface cannot be obtained.
            Exception: If unable to set property value
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("obj", "prop_set", "name", "value")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_property() got an unexpected keyword argument")
            keys = ("obj", "prop_set")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            ka[2] = kwargs.get("name", None)
            ka[3] = kwargs.get("value", None)
            return ka

        if count != 3:
            raise TypeError("set_property() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mLo.Lo.is_uno_interfaces(kargs[1], XPropertySet):
            prop_set = cast(XPropertySet, kargs[1])
        else:
            prop_set = mLo.Lo.qi(XPropertySet, kargs[1])
            if prop_set is None:
                raise mEx.MissingInterfaceError(XPropertySet)
        try:
            prop_set.setPropertyValue(kargs[2], kargs[3])
        except Exception as e:
            raise Exception(f"Could not set property '{kargs[2]}'") from e

    # endregion set_property()

    # region    set_properties()
    @overload
    @classmethod
    def set_properties(cls, obj: object, names: Sequence[str], vals: Sequence[object]) -> None:
        """
        Set Properties

        Args:
            obj (object): Object that implements XPropertySet interface
            names (Sequence[str]): Property Names
            vals (Sequence[object]): Property Values

        Raises:
            MissingInterfaceError: if obj does not implement XPropertySet interface
            MultiError: If unable to set a property
        """
        ...

    @overload
    @classmethod
    def set_properties(cls, prop_set: XPropertySet, names: Sequence[str], vals: Sequence[object]) -> None:
        """
        Set Properties

        Args:
            prop_set (XPropertySet): Property Set
            names (Sequence[str]): Property Names
            vals (Sequence[object]): Property Values

        Raises:
            MultiError: If unable to set a property
        """
        ...

    @overload
    @classmethod
    def set_properties(cls, obj: object, from_obj: object) -> None:
        """
        Set properties

        Properties of ``from_obj`` are assigned to ``obj``

        Args:
            obj (object): Object that implements XPropertySet interface
            from_obj (object): Other object that implements XPropertySet interface

        Raises:
            MissingInterfaceError: if obj does not implement XPropertySet interface
            MissingInterfaceError: if from_obj does not implement XPropertySet interface
            MultiError: If unable to set a property
        """
        ...

    @overload
    @classmethod
    def set_properties(cls, prop_set: XPropertySet, from_props: XPropertySet) -> None:
        """
        Set properties

        Properties of ``from_props`` are assigned to ``prop_set``

        Args:
            prop_set (XPropertySet): Property set
            from_props (XPropertySet): Property set

        Raises:
            MultiError: If unable to set a property
        """
        ...

    @classmethod
    def set_properties(cls, *args, **kwargs) -> None:
        """
        Set Properties

        Args:
            obj (object): Object that implements XPropertySet interface
            prop_set (XPropertySet): Property set
            from_props (XPropertySet): Property set
            from_obj (object): Other object that implements XPropertySet interface
            names (Sequence[str]): Property Names
            vals (Sequence[object]): Property Values

        Raises:
            MissingInterfaceError: if obj does not implement XPropertySet interface
            MissingInterfaceError: if from_obj does not implement XPropertySet interface
            MultiError: If unable to set a property
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("obj", "prop_set", "names", "from_obj", "from_props", "vals")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_properties() got an unexpected keyword argument")
            keys = ("obj", "prop_set")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            keys = ("names", "from_obj", "from_props")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("vals", None)
            return ka

        if not count in (2, 3):
            raise TypeError("set_properties() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mInfo.Info.is_type_interface(kargs[1], XPropertySet.__pyunointerface__):
            # set_properties(cls, prop_set: XPropertySet, names: Sequence[str], vals: Sequence[object]) -> None
            # set_properties(cls, prop_set: XPropertySet, from_props: XPropertySet) -> None
            prop_set = cast(XPropertySet, kargs[1])
        else:
            # set_properties(cls, obj: object, names: Sequence[str], vals: Sequence[object]) -> None
            # set_properties(cls, obj: object, from_obj: object) -> None
            prop_set = mLo.Lo.qi(XPropertySet, kargs[1])
            if prop_set is None:
                raise mEx.MissingInterfaceError(XPropertySet)

        if count == 3:
            # set_properties(cls, prop_set: XPropertySet, names: Sequence[str], vals: Sequence[object]) -> None
            # set_properties(cls, obj: object, names: Sequence[str], vals: Sequence[object]) -> None
            cls._set_properties_by_vals(prop_set=prop_set, names=kargs[2], vals=kargs[3])
            return

        elif count == 2:
            if mInfo.Info.is_type_interface(kargs[2], XPropertySet.__pyunointerface__):
                # set_properties(cls, prop_set: XPropertySet, from_props: XPropertySet) -> None
                from_props = cast(XPropertySet, kargs[2])
            else:
                # set_properties(cls, obj: object, from_obj: object) -> None
                from_props = mLo.Lo.qi(XPropertySet, kargs[1])
                if from_props is None:
                    raise mEx.MissingInterfaceError(XPropertySet)
            cls._set_properties_from_props(prop_set=prop_set, from_props=from_props)

    @classmethod
    def _set_properties_by_vals(cls, prop_set: XPropertySet, names: Sequence[str], vals: Sequence[object]) -> None:
        errs = []
        for i, name in enumerate(names):
            try:
                prop_set.setPropertyValue(name, vals[i])
            except PropertyVetoException as e:
                errs.append(mEx.PropertyError(f"Could not set readonly-property {name}: {e}", e))
            except Exception as e:
                errs.append(Exception(f"Cound ont set property {name}: {e}", e))
        if len(errs) > 0:
            raise mEx.MultiError(errs)

    @classmethod
    def _set_properties_from_props(cls, prop_set: XPropertySet, from_props: XPropertySet) -> None:
        errs = []
        if prop_set is None:
            print(f"Property set is null; cannot set properties")
            return
        if from_props is None:
            print("Source property set is null; cannot set properties")
            return
        nms = cls.get_prop_names(from_props)
        for itm in nms:
            try:
                prop_set.setPropertyValue(itm, cls.get_property(from_props, itm))
            except Exception as e:
                errs.append(Exception(f"Could not set property '{itm}': {e}", e))
        if len(errs) > 0:
            raise mEx.MultiError(errs)

    # endregion set_properties()
    # endregion ---------------- set properties -----------------------

    # region ------------------- get properties ------------------------

    # region    get_property()
    @overload
    @staticmethod
    def get_property(obj: object, name: str) -> object:
        """
        Gets a property value from property set

        Args:
            obj (object): object that implements XPropertySet interface
            xprops (XPropertySet): property set
            name (str): property name

        Raises:
            PropertyNotFoundError: If unable to get property
            MissingInterfaceError: if obj does not implement XPropertySet interface

        Returns:
            object: property value
        """
        ...

    @overload
    @staticmethod
    def get_property(prop_set: XPropertySet, name: str) -> object:
        """
        Gets a property value from property set

        Args:
            obj (object): object that implements XPropertySet interface
            xprops (XPropertySet): property set
            name (str): property name

        Raises:
            PropertyNotFoundError: If unable to get property

        Returns:
            object: property value
        """
        ...

    @staticmethod
    def get_property(*args, **kwargs) -> object:
        """
        Gets a property value from property set

        Args:
            obj (object): object that implements XPropertySet interface
            xprops (XPropertySet): property set
            name (str): property name

        Raises:
            PropertyNotFoundError: If unable to get property
            MissingInterfaceError: if obj does not implement XPropertySet interface

        Returns:
            object: property value
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("obj", "prop_set", "name")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_property() got an unexpected keyword argument")
            keys = ("obj", "prop_set")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            ka[2] = kwargs.get("name", None)
            return ka

        if count != 2:
            raise TypeError("get_property() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mInfo.Info.is_type_interface(kargs[1], XPropertySet.__pyunointerface__):
            prop_set = cast(XPropertySet, kargs[1])
        else:
            prop_set = mLo.Lo.qi(XPropertySet, kargs[1])
            if prop_set is None:
                raise mEx.MissingInterfaceError(XPropertySet)
        name = kargs[2]
        try:
            try:
                return prop_set.getPropertyValue(name)
            except RuntimeException as e:
                mEx.PropertyError(name, f"Could not get runtime property '{name}': {e}")
        except Exception as e:
            raise mEx.PropertyNotFoundError(prop_name=name) from e

    # endregion    get_property()

    @staticmethod
    def get_properties(obj: object) -> Tuple[Property, ...]:
        """
        Get properties

        Args:
            obj (object): Object that implements XPropertySet interface

        Raises:
            MissingInterfaceError: if obj dos not implement XPropertySet

        Returns:
            Tuple[Property, ...]: Properties of obj
        """
        prop_set = mLo.Lo.qi(XPropertySet, obj)
        if prop_set is None:
            raise mEx.MissingInterfaceError(XPropertySet)
        props = list(prop_set.getPropertySetInfo().getProperties())
        props.sort(key=lambda prop: prop.Name)
        return tuple(props)

    @staticmethod
    def get_prop_names(obj: object) -> Tuple[str, ...]:
        """
        Gets property names

        Args:
            obj (object): Object that implements XPropertySet interface

        Raises:
            MissingInterfaceError: if obj dos not implement XPropertySet

        Returns:
            Tuple[str, ...]: Property names of obj
        """
        prop_set = mLo.Lo.qi(XPropertySet, obj)
        if prop_set is None:
            raise mEx.MissingInterfaceError(XPropertySet)
        props = prop_set.getPropertySetInfo().getProperties()
        nms = []
        for prop in props:
            nms.append(prop.Name)
        return tuple(nms)

    @staticmethod
    def get_value(name: str, props: Iterable[PropertyValue]) -> object:
        """
        Get a property value from properties

        Args:
            name (str): property name
            props (Iterable[PropertyValue]): Properties to search

        Raises:
            PropertyError: If property name is not found.

        Returns:
            object: Property Value
        """
        for prop in props:
            if prop.Name == name:
                return prop.Value
        raise mEx.PropertyNotFoundError(
            name,
        )

    # endregion ---------------- get properties ------------------------

    # region ------------------- show properties array -----------------

    @classmethod
    def show_indexed_props(cls, title: str, obj: object) -> None:
        """
        Prints objects properties to console

        Args:
            title (str): Title to print
            obj (object): object that implements XIndexAccess

        Example:
            .. code-block:: python

                >>> from ooodev.office.calc import Calc
                >>> from ooodev.utils.props import Props
                >>> from ooodev.utils.lo import Lo
                >>> loader = Lo.load_office()
                >>> doc = Calc.create_doc(loader)
                >>> sheets = Calc.get_sheets(doc)
                >>> Props.show_indexed_props("Sheets Property", sheets)
                Indexed Properties for 'Sheets Property':
                No. of elements: 1
                Elem 0 Properties
                  AbsoluteName: $Sheet1.$A$1:$AMJ$1048576
                  AsianVerticalMode: False
                  AutomaticPrintArea: True
                  BorderColor: None
                  BottomBorder: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
                  BottomBorder2: (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
                  CellBackColor: -1
                  CellProtection: (com.sun.star.util.CellProtection){ IsLocked = (boolean)true, IsFormulaHidden = (boolean)false, IsHidden = (boolean)false, IsPrintHidden = (boolean)false }
                  CellStyle: Default
                  CharColor: -1
                  ...
        """
        print(f"Indexed Properties for '{title}':")
        in_acc = mLo.Lo.qi(XIndexAccess, obj)
        if in_acc is None:
            print("Could not convert object to an IndexAccess container")
            return

        num_elems = in_acc.getCount()
        print(f"No. of elements: {num_elems}")
        if num_elems == 0:
            return
        for i in range(num_elems):
            try:
                # PropertyValue[] props = Lo.qi(PropertyValue[].class, inAcc.getByIndex(i));
                # above line is original java code.
                # perhapsh thre is a way to alos include PropertyValue[] queryInterface
                props = mLo.Lo.qi(XPropertySet, in_acc.getByIndex(i))
                if props is None:
                    return

                cls.show_props(f"Elem {i}", props)
                print("----")
            except Exception as e:
                print(f"Could not get elem {i}: {e}")

    @classmethod
    def prop_value_to_string(cls, val: object) -> str:
        """
        Gets property values a a string

        Args:
            val (object): Values such as a iterable of iterable or
                object that implements XPropertySet or
                a string

        Returns:
            str: A string representing properties

        Example:
             .. code-block:: python

                >>> from ooodev.office.calc import Calc
                >>> from ooodev.utils.props import Props
                >>> from ooodev.utils.lo import Lo
                >>> loader = Lo.load_office()
                >>> doc = Calc.create_doc(loader)
                >>> sheet = Calc.get_sheet(doc=doc, index=0)
                >>> prop_str = Props.prop_value_to_string(sheet)
                >>> print(prop_str)
                [
                    AbsoluteName = $Sheet1.$A$1:$AMJ$1048576
                    AsianVerticalMode = False
                    AutomaticPrintArea = True
                    BorderColor = None
                    BottomBorder = (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
                    BottomBorder2 = (com.sun.star.table.BorderLine2){ (com.sun.star.table.BorderLine){ Color = (long)0x0, InnerLineWidth = (short)0x0, OuterLineWidth = (short)0x0, LineDistance = (short)0x0 }, LineStyle = (short)0x0, LineWidth = (unsigned long)0x0 }
                    CellBackColor = -1
                    CellProtection = (com.sun.star.util.CellProtection){ IsLocked = (boolean)true, IsFormulaHidden = (boolean)false, IsHidden = (boolean)false, IsPrintHidden = (boolean)false }
                    CellStyle = Default
                    CharColor = -1
                    ...
                ]
        """

        def get_pv_str(vals) -> str:
            lines = []
            for p in vals:
                try:
                    # value = cls.get_property()
                    lines.append(f"{p.Name} = {p.Value}")
                except AttributeError:
                    continue
            return "[\n    " + "\n    ".join(lines) + "\n]"

        def get_property_set_str(prop_set, props) -> str:
            lines = []
            for p in props:
                value = cls.get_property(prop_set, p.Name)
                lines.append(f"{p.Name} = {value}")
            return "[\n    " + "\n    ".join(lines) + "\n]"

        if val is None:
            return ""
        if isinstance(val, str):
            return val
        try:
            _ = iter(val)
        except TypeError:
            # not iterable
            is_iter = False
        else:
            # iterable
            is_iter = True
        if is_iter:

            try:
                if isinstance(val[0], str):
                    # assume Iterable[str]
                    return ", ".join(val)
            except Exception:
                pass

            # assume Iterable[PropertyValue]
            return get_pv_str(val)
        else:
            xprops = mLo.Lo.qi(XPropertySet, val)
            if xprops is not None:
                lst = list(cls.props_set_to_tuple(xprops))
                lst.sort(key=lambda prop: prop.Name)
                return get_property_set_str(xprops, lst)
            return str(val)

    @classmethod
    def show_values(cls, name: str, props: Iterable[PropertyValue]) -> None:
        """
        Prints property to console

        Args:
            name (str): Property Name
            props (Iterable[PropertyValue]): Properties
        """
        for prop in props:
            if prop.Name == name:
                print(f"{prop.Name}: {cls.prop_value_to_string(prop.Value)}")
                return
        print(f"{name} not found")

    # endregion ---------------- show properties array -----------------

    # region ------------------- show properties of an Object ----------

    @classmethod
    def show_obj_props(cls, prop_kind: str, obj: object) -> None:
        """
        Prints properties for an object to the console

        Args:
            prop_kind (str): The kind of properties that is displayed in console
            obj (object): object the implements XPropertySet interface.
        """
        prop_set = mLo.Lo.qi(XPropertySet, obj)
        if prop_set is None:
            print(f"no {prop_kind} properties found")
            return
        cls.show_props(prop_kind=prop_kind, props_set=prop_set)

    # region    show_props()
    @overload
    @classmethod
    def show_props(cls, title: str, props: Sequence[PropertyValue]) -> None:
        """
        Prints properties to console

        Args:
            title (str): Title to use that is displayed in console
            props (Sequence[PropertyValue]): Properties to print
        """
        ...

    @overload
    @classmethod
    def show_props(cls, prop_kind: str, props_set: XPropertySet) -> None:
        """
        Prints properties to console

        Args:
            prop_kind (str): The kind of properties that is displayed in console
            props_set (XPropertySet): object the implements XPropertySet interface.
        """
        ...

    @classmethod
    def show_props(cls, *args, **kwargs) -> None:
        """
        Prints properties to console

        Args:
            title (str): Title to use that is displayed in console
            props (Sequence[PropertyValue]): Properties to print
            prop_kind (str): The kind of properties that is displayed in console
            props_set (XPropertySet): object the implements XPropertySet interface.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("title", "props", "prop_kind", "props_set")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("show_props() got an unexpected keyword argument")
            keys = ("title", "prop_kind")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            keys = ("props", "props_set")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("show_props() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mInfo.Info.is_type_interface(kargs[2], "com.sun.star.beans.XPropertySet"):
            # def show_props(prop_kind: str, props_set: XPropertySet)
            return cls._show_props_str_xpropertyset(prop_kind=kargs[1], props_set=kargs[2])
        else:
            # def show_props(title: str, props: Sequence[PropertyValue])
            return cls._show_props_str_props(title=kargs[1], props=kargs[2])

    @classmethod
    def _show_props_str_props(cls, title: str, props: Sequence[PropertyValue]) -> None:
        print(f"Properties for '{title}':")
        if props is None:
            print("  none found")
            return
        lst = list(props)
        lst.sort(key=lambda prop: prop.Name)
        for prop in lst:

            print(f"  {prop.Name}: {cls.prop_value_to_string(prop.Value)}")
        print()

    @classmethod
    def _show_props_str_xpropertyset(cls, prop_kind: str, props_set: XPropertySet) -> None:
        props = cls.props_set_to_tuple(props_set)
        if props is None:
            print(f"No {prop_kind} properties found")
            return
        lst = list(props)
        lst.sort(key=lambda prop: prop.Name)
        print(f"{prop_kind} Properties")
        for prop in lst:
            prop_value = cls.get_property(props_set, prop.Name)
            print(f"  {prop.Name}: {prop_value}")
        print()

    # endregion  show_props()

    @staticmethod
    def props_set_to_tuple(xprops: XPropertySet) -> Tuple[Property, ...]:
        """
        Converts Property Set to Tuple of Property

        Args:
            xprops (XPropertySet): Property set

        Returns:
            Tuple[Property, ...]: Tuple of Property
        """
        if xprops is None:
            return ()
        xprops_info = xprops.getPropertySetInfo()
        return xprops_info.getProperties()

    props_set_to_array = props_set_to_tuple

    @staticmethod
    def show_property(p: Property) -> str:
        """
        Gets a property Name and type as string.

        Args:
            p (Property): Property to print

        Returns:
            str: Property inf format of 'Name: TypeName'
        """
        # p.Type is uno.Type
        return f"{p.Name}: {p.Type.typeName}"

    # endregion ---------------- show properties of an Object ----------

    # region ------------------- others --------------------------------
    @staticmethod
    def has_property(obj: object, name: str) -> bool:
        """
        Gets if a object contains a property matching name

        Args:
            obj (object): An object that implements XPropertySet
            name (str): Property Name

        Returns:
            bool: True if obj contains Property that matches name; Otherwise, False
        """
        prop_set = mLo.Lo.qi(XPropertySet, obj)
        if prop_set is None:
            return False
        return prop_set.getPropertySetInfo().hasPropertyByName(name)

    @classmethod
    def show_doc_type_props(cls, type: str) -> None:
        """
        Prints doc type info to console

        Args:
            type (str): doc type
        """
        if type is None:
            print("type is None")
            return
        xtype_detect = mLo.Lo.create_instance_mcf(XTypeDetection, "com.sun.star.document.TypeDetection")
        if xtype_detect is None:
            print("No type detector reference")
            return

        xname_access = mLo.Lo.qi(XNameAccess, xtype_detect)
        try:
            props = xname_access.getByName(type)
            cls.show_props(type, props)
        except Exception:
            print(f"No properties for '{type}'")

    @staticmethod
    def get_bound_props(props_set: XMultiPropertySet) -> Tuple[str, ...]:
        """
        Gets bound Properties

        Args:
            props_set (XMultiPropertySet): Multi Property Set

        Returns:
            Tuple[str, ...]: Bound names
        """
        props = props_set.getPropertySetInfo().getProperties()
        names = []
        for p in props:
            is_writable = (p.Attributes & PropertyAttribute.READONLY) == 0
            is_not_null = (p.Attributes & PropertyAttribute.MAYBEVOID) == 0
            is_bound = (p.Attributes & PropertyAttribute.BOUND) == 0
            if is_writable and is_not_null and is_bound:
                names.append(p.Name)
        names_len = len(names)
        if names_len == 0:
            print("No suitable properties were found")
            return ()
        print(f"No. of suitable properties: {names_len}")
        names.sort()
        return tuple(names)

    # endregion ---------------- others --------------------------------
