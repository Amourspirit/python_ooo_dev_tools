# coding: utf-8
"""make/get/set properties in an array"""

# Python conversion of Props.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
# region Imports
from __future__ import annotations
import contextlib
from typing import Any, Dict, List, Iterable, Optional, Sequence, Tuple, TYPE_CHECKING, cast, overload, Union
import uno

from com.sun.star.beans import PropertyAttribute  # const
from com.sun.star.beans import PropertyVetoException
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.beans import XFastPropertySet
from com.sun.star.beans import XPropertySet
from com.sun.star.beans import XPropertyState
from com.sun.star.container import XIndexAccess
from com.sun.star.container import XNameAccess
from com.sun.star.document import XTypeDetection
from com.sun.star.ui import ItemStyle  # const
from com.sun.star.ui import ItemType  # const

from ooo.dyn.beans.property_value import PropertyValue
from ooo.dyn.beans.named_value import NamedValue
from ooo.dyn.beans.property import Property
from ooo.dyn.beans.string_pair import StringPair


from ooodev.utils import gen_util as gUtil
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.events.args.key_val_args import KeyValArgs
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.event_singleton import _Events
from ooodev.events.props_named_event import PropsNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.utils.helper.dot_dict import DotDict


if TYPE_CHECKING:
    from com.sun.star.beans import XPropertySetInfo
    from com.sun.star.beans import XMultiPropertySet

    # import module and not module content to avoid circular import issue.
    # https://stackoverflow.com/questions/22187279/python-circular-importing
    #
    # using lazy loading: https://snarky.ca/lazy-importing-in-python-3-7/
# endregion Imports


class Props:
    """make/get/set properties in an array"""

    # region ------------------- make properties -----------------------
    @staticmethod
    def make_prop_value(name: Optional[str] = None, value: Optional[Any] = None) -> PropertyValue:
        """
        Makes a Uno Property Value and assigns name and value if present.

        |lo_safe|

        Args:
            name (Optional[str], optional): Property name
            value (Optional[Any], optional): Property value

        Returns:
            PropertyValue: com.sun.star.beans.PropertyValue

        See Also:
            `LibreOffice API PropertyValue <https://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1beans_1_1PropertyValue.html>`_
        """
        p = cast(PropertyValue, uno.createUnoStruct("com.sun.star.beans.PropertyValue"))
        if name is not None:
            p.Name = name
        if value is not None:
            p.Value = value
        return p

    @staticmethod
    def make_sting_pair(first: str = "", second: str = "") -> StringPair:
        """
        Makes a Uno String Pair and assigns first and second.

        |lo_safe|

        Args:
            first (str, optional): First string.
            second (str, optional): Second string.

        Returns:
            StringPair: com.sun.star.beans.StringPair

        .. versionadded:: 0.40.0
        """
        p = cast(StringPair, uno.createUnoStruct("com.sun.star.beans.StringPair"))
        p.First = first
        p.Second = second
        return p

    @classmethod
    def make_bar_item(
        cls, cmd: str, item_name: str
    ) -> Tuple[PropertyValue, PropertyValue, PropertyValue, PropertyValue, PropertyValue]:
        """
        Properties for a toolbar item using a name and an image.

        problem: image does not appear next to text on toolbar

        |lo_safe|

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
    def make_props(cls, **kwargs: Any) -> Tuple[PropertyValue, ...]:
        """
        Make Properties.

        |lo_safe|

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
        lst = [cls.make_prop_value(name=k, value=v) for k, v in kwargs.items()]
        return tuple(lst)

    @classmethod
    def make_strings(cls, **kwargs: str) -> Tuple[StringPair, ...]:
        """
        Make String Pairs.

        |lo_safe|

        Keyword Args:
            kwargs (Dict[str, str]): Each key, value pair is assigned to a StringPair.

        Returns:
            Tuple[StringPair, ...]: Tuple of String Pairs

        .. versionadded:: 0.40.0
        """
        lst = [cls.make_sting_pair(first=k, second=v) for k, v in kwargs.items()]
        return tuple(lst)

    @classmethod
    def make_props_any(cls, **kwargs: Any) -> Any:
        """
        Makes a uno.Any object for properties.

        Keyword Args:
            kwargs (Dict[str, Any]): Each key, value pair is assigned to a PropertyValue.

        Returns:
            uno.Any: Array of ``[]com.sun.star.beans.PropertyValue``

        .. versionadded:: 0.40.0
        """
        props = cls.make_props(**kwargs)
        return uno.Any("[]com.sun.star.beans.PropertyValue", props)  # type: ignore

    # endregion ---------------- make properties -----------------------

    # region ------------------- Data -----------------------------------

    @staticmethod
    def data_to_dict(
        data: Sequence[Tuple[str, Any]] | Sequence[List[str]] | Sequence[PropertyValue] | Sequence[NamedValue]
    ) -> Dict[str, Any]:
        """
        Convert tuples, list, PropertyValue, NamedValue to dictionary.

        |lo_safe|

        Args:
            data (Sequence[Tuple[str, Any]] | Sequence[List[str]] | Sequence[PropertyValue] | Sequence[NamedValue]): Data to convert to dictionary.
                If data is a sequence of tuples or list, the first element is the key and the second element is the value.

        Returns:
            Dict[str, Any]: Dictionary.

        .. versionadded:: 0.40.0
        """
        d = {}
        if not data:
            return d

        if isinstance(data[0], (tuple, list)):
            data_seq = cast(Union[Sequence[Tuple[str, Any]], Sequence[List[str]]], data)
            d = {r[0]: r[1] for r in data_seq}
        elif isinstance(data[0], (PropertyValue, NamedValue)):
            data_seq = cast(Union[Sequence[PropertyValue], Sequence[NamedValue]], data)
            d = {r.Name: r.Value for r in data_seq}

        return d

    # endregion ---------------- Data -----------------------------------

    # region DotDict
    @staticmethod
    def props_to_dot_dict(props: Sequence[PropertyValue]) -> DotDict:
        """
        Convert a sequence of PropertyValues to a DotDict.

        |lo_safe|

        Args:
            props (Sequence[PropertyValue]): Properties to convert to DotDict.

        Returns:
            DotDict: DotDict

        .. versionadded:: 0.45.0
        """
        return DotDict(**{prop.Name: prop.Value for prop in props})

    @staticmethod
    def dot_dict_to_props(dot_dict: DotDict) -> List[PropertyValue]:
        """
        Convert a DotDict to a list of PropertyValues.

        |lo_safe|

        Args:
            dot_dict (DotDict): DotDict to convert to PropertyValues.

        Returns:
            List[PropertyValue]: List of PropertyValues

        .. versionadded:: 0.45.0
        """
        result = []
        for k, v in dot_dict.items():
            pv = PropertyValue()
            pv.Name = k
            pv.Value = v
            result.append(pv)
        return result

    # endregion DotDict

    # region ------------------- uno -----------------------------------
    @staticmethod
    def any(*elements: object) -> Any:  # type: ignore
        """
        Gets a uno.Any object for elements.

        |lo_safe|

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
        if not elements:
            raise ValueError("No args to create unn.Any object")
        obj = elements[0]
        if isinstance(obj, uno.Type):
            type_name = obj.typeName
            return uno.Any(obj, [*elements])  # type: ignore
        elif isinstance(obj, str):
            type_name = "string"
        elif isinstance(obj, int):
            type_name = "short"
        else:
            type_name = mInfo.Info.get_type_name(obj)
        if type_name is None:
            raise ValueError("Unable to get type name to create uno.Any object")
        return uno.Any(f"[]{type_name}", [*elements])  # type: ignore

    # endregion ---------------- uno -----------------------------------

    # region ------------------- set properties ------------------------
    @staticmethod
    def set_prop(props: Iterable[PropertyValue], name: str, value: object) -> bool:
        """
        Sets property value for the first property that has matching name.

        |lo_safe|

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
    def get_xproperty_set(obj: object) -> XPropertySet:
        """
        Gets Property Set.

        |lo_safe|

        Args:
            obj (object): object that implements ``XPropertySet``

        Raises:
            MissingInterfaceError: if ``obj`` does not implement ``XPropertySet``.

        Returns:
            XPropertySet: Property Set
        """
        return mLo.Lo.qi(XPropertySet, obj, True)

    @staticmethod
    def get_xproperty_set_fast(obj: object) -> XFastPropertySet:
        """
        Gets Fast Property Set.

        |lo_safe|

        Args:
            obj (object): object that implements ``XFastPropertySet``

        Raises:
            MissingInterfaceError: if ``obj`` does not implement ``XFastPropertySet``.

        Returns:
            XFastPropertySet: Property Set
        """
        return mLo.Lo.qi(XFastPropertySet, obj, True)

    # region get_xproperty_fast_value()
    @overload
    @classmethod
    def get_xproperty_fast_value(cls, *, handle: int, fps: XFastPropertySet) -> Any: ...

    @overload
    @classmethod
    def get_xproperty_fast_value(cls, *, prop: Property, fps: XFastPropertySet) -> Any: ...

    @overload
    @classmethod
    def get_xproperty_fast_value(cls, *, handle: int, obj: object) -> Any: ...

    @classmethod
    def get_xproperty_fast_value(cls, *args, **kwargs) -> Any:
        """
        Gets a fast property value via ``XFastPropertySet``

        |lo_safe|

        Arguments:
            handle (int): handle of the property.
            fps (XFastPropertySet): property set that contains the property for matching handle.
            prop (Property): Property that handle can be obtained from.
            obj (object): an object that ``XFastPropertySet`` can be obtained from.

        Raises:
            PropertyGeneralError: If an error occurs.

        Returns:
            Any: Property Value
        """
        if args:
            raise TypeError(f"get_xproperty_fast_value() takes 0 positional arguments but {len(args)} was given")

        count = len(kwargs)

        if count != 2:
            raise TypeError("get_xproperty_fast_value() got an invalid number of arguments")

        try:
            handle = cast(int, kwargs.get("handle", None))
            if handle is None:
                handle = cast(Property, kwargs.get("prop")).Handle

            fps = cast(XFastPropertySet, kwargs.get("fps", None))
            if fps:
                return fps.getFastPropertyValue(handle)

            fps = cls.get_xproperty_set_fast(kwargs.get("obj"))
            return fps.getFastPropertyValue(handle)
        except Exception as e:
            raise mEx.PropertyGeneralError("Error getting fast property value") from e

    # endregion get_xproperty_fast_value()

    @classmethod
    def get_property_set_info(cls, obj: object) -> XPropertySetInfo:
        """
        Gets property set info.

        |lo_safe|

        Args:
            obj (object): Object that implements ``XPropertySet``.

        Raises:
            PropertyGeneralError: If error occurs

        Returns:
            XPropertySetInfo: Property Set Info
        """
        try:
            x_set = cls.get_xproperty_set(obj)
            result = x_set.getPropertySetInfo()
            if result is None:
                raise mEx.NoneError("None Value: x_set.getPropertySetInfo() returned None")
            return result
        except Exception as e:
            raise mEx.PropertyGeneralError("Error getting property set info") from e

    @classmethod
    def get_property_by_name(cls, obj: object, name: str) -> Property:
        """
        Gets a property by name.

        |lo_safe|

        Property instances do not contain a value.

        Args:
            obj (object): Object to get a property from
            name (str): Property name to lookup in ``obj``.

        Raises:
            PropertyError: If error occurs

        Returns:
            Property: Property instance.
        """
        try:
            info = cls.get_property_set_info(obj)
            return info.getPropertyByName(name)  # type: ignore
        except Exception as e:
            raise mEx.PropertyError(name) from e

    @classmethod
    def has_property(cls, obj: object, name: str) -> bool:
        """
        Gets is an object contains a Property.

        |lo_safe|

        Args:
            obj (object): object that implements ``XPropertySet``.
            name (str): Property Name.

        Returns:
            bool: ``True`` if object contains property with name; Otherwise, ``False``.
        """
        with contextlib.suppress(Exception):
            x_set = cls.get_property_set_info(obj)
            return x_set.hasPropertyByName(name)
        return False

    @staticmethod
    def get_prop(props: Iterable[PropertyValue], name: str) -> object | None:
        """
        Gets property value for property that matches name.

        |lo_safe|

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
    @classmethod
    def set_property(cls, obj: object, name: str, value: object) -> None: ...

    @overload
    @classmethod
    def set_property(cls, prop_set: XPropertySet, name: str, value: object) -> None: ...

    @classmethod
    def set_property(cls, *args, **kwargs) -> None:
        """
        Sets the value of the property with the specified name.

        |lo_safe|

        Args:
            obj (object): object that implements XPropertySet interface
            prop_set (XPropertySet): Property set
            name (str): Name of property to set value of
            value (object): property value

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_SETTING` :eventref:`src-docs-props-event-setting`
                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_SET` :eventref:`src-docs-props-event-set`

        Raises:
            MissingInterfaceError: If required interface cannot be obtained.
            PropertySetError: If unable to set property value

        Returns:
            None:

        Note:
            If a Event is canceled then property is not set. No error occurs.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("obj", "prop_set", "name", "value")
            check = all(key in valid_keys for key in kwargs)
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
            raise TypeError("set_property() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        try:
            cls.set(kargs[1], **{kargs[2]: kargs[3]})
        except mEx.MissingInterfaceError:
            raise
        except mEx.MultiError as me:
            err = me.errors[0]
            if isinstance(err, (mEx.PropertyNotFoundError, mEx.PropertySetError)):
                raise err
            raise mEx.PropertySetError(f"Could not set property '{kargs[2]}'") from err

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

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_SETTING` :eventref:`src-docs-props-event-setting`
                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_SET` :eventref:`src-docs-props-event-set`

        Note:
            If a Event is canceled then that property is not set. No error occurs.

            If ``MultiError`` occurs only the properties that raised an error is part of the error object.
            The remaining properties will still be set.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("obj", "prop_set", "names", "from_obj", "from_props", "vals")
            check = all(key in valid_keys for key in kwargs)
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

        if count not in (2, 3):
            raise TypeError("set_properties() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mInfo.Info.is_type_interface(kargs[1], "com.sun.star.beans.XPropertySet"):
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
            if mInfo.Info.is_type_interface(kargs[2], "com.sun.star.beans.XPropertySet"):
                # set_properties(cls, prop_set: XPropertySet, from_props: XPropertySet) -> None
                from_props = cast(XPropertySet, kargs[2])
            else:
                # set_properties(cls, obj: object, from_obj: object) -> None
                from_props = mLo.Lo.qi(XPropertySet, kargs[1], True)
            cls._set_properties_from_props(prop_set=prop_set, from_props=from_props)

    @classmethod
    def _set_properties_by_vals(cls, prop_set: XPropertySet, names: Sequence[str], vals: Sequence[object]) -> None:
        errs = []
        for i, name in enumerate(names):
            has_error = False
            cargs = None
            try:
                cargs = KeyValCancelArgs(Props.set_properties.__qualname__, name, vals[i])
                cargs.event_data = prop_set
                _Events().trigger(PropsNamedEvent.PROP_SETTING, cargs)
                if cargs.cancel:
                    continue
                prop_set.setPropertyValue(cargs.key, cargs.value)
            except PropertyVetoException as e:
                has_error = True
                errs.append(mEx.PropertySetError(f'Could not set readonly-property "{name}"', e))
            except UnknownPropertyException as e:
                has_error = True
                errs.append(mEx.PropertyNotFoundError(name, e))
            except Exception as e:
                has_error = True
                errs.append(Exception(f'Could not set property "{name}"', e))
            if not has_error and cargs:
                _Events().trigger(PropsNamedEvent.PROP_SET, KeyValArgs.from_args(cargs))  # type: ignore
        if errs:
            raise mEx.MultiError(errs)

    @classmethod
    def _set_properties_from_props(cls, prop_set: XPropertySet, from_props: XPropertySet) -> None:
        errs = []
        if prop_set is None:
            print("Property set is null; cannot set properties")
            return
        if from_props is None:
            print("Source property set is null; cannot set properties")
            return
        nms = cls.get_prop_names(from_props)
        for itm in nms:
            has_error = False
            cargs = None
            try:
                cargs = KeyValCancelArgs(Props.set_properties.__qualname__, itm, cls.get_property(from_props, itm))
                cargs.event_data = prop_set
                _Events().trigger(PropsNamedEvent.PROP_SETTING, cargs)
                if cargs.cancel:
                    continue
                prop_set.setPropertyValue(cargs.key, cargs.value)
            except PropertyVetoException as e:
                has_error = True
                errs.append(mEx.PropertySetError(f'Could not set readonly-property "{itm}"', e))
            except UnknownPropertyException as e:
                has_error = True
                errs.append(mEx.PropertyNotFoundError(itm, e))
            except Exception as e:
                has_error = True
                errs.append(Exception(f'Could not set property "{itm}"', e))
            if not has_error and cargs:
                _Events().trigger(PropsNamedEvent.PROP_SET, KeyValArgs.from_args(cargs))  # type: ignore
        if errs:
            raise mEx.MultiError(errs)

    # endregion set_properties()

    @staticmethod
    def _set_by_attribute(obj: object, name: str, value: Any) -> bool:
        """Lo Safe Method"""
        # there is a bug in LibreOffice that in some cases getting XPropertySet
        # returns the interface but it is missing setPropertyValue and or getPropertySetInfo
        # this is the case with XChartDocument.getPageBackground() and XChartDocument.getFirstDiagram().getWall()
        # See Chart2.set_background_colors()
        #
        # the same issue (ct is XChartType) with:
        #   white_day_ps = mLo.Lo.qi(XPropertySet, mProps.Props.get(ct, "WhiteDay"), True)
        #   black_day_ps = mLo.Lo.qi(XPropertySet, mProps.Props.get(ct, "BlackDay"), True)
        # see Chart2.color_stock_bars()
        # Even though the interface is missing setPropertyValue and or getPropertySetInfo the
        # properties are still there.
        # That is why this method. It is a fallback when bug is found to set value by attribute.
        # see Props.set()

        if not hasattr(obj, name):
            return False
        with contextlib.suppress(AttributeError):
            setattr(obj, name, value)
            return True
        return False

    @staticmethod
    def _get_by_attribute(obj: object, name: str) -> Tuple[bool, Any]:
        """LO Safe Method"""
        # there is a bug in LibreOffice that in some cases getting XPropertySet
        # returns the interface but it is missing setPropertyValue and or getPropertySetInfo
        # this is the case with XChartDocument.getPageBackground() and XChartDocument.getFirstDiagram().getWall()
        # See Chart2.set_background_colors()
        #
        # the same issue (ct is XChartType) with:
        #   white_day_ps = mLo.Lo.qi(XPropertySet, mProps.Props.get(ct, "WhiteDay"), True)
        #   black_day_ps = mLo.Lo.qi(XPropertySet, mProps.Props.get(ct, "BlackDay"), True)
        # see Chart2.color_stock_bars()
        # Even though the interface is missing setPropertyValue and or getPropertySetInfo the
        # properties are still there.
        # That is why this method. It is a fallback when bug is found to get value by attribute.
        # see Props.get()
        with contextlib.suppress(AttributeError):
            return (True, getattr(obj, name))
        return (False, None)

    @classmethod
    def set(cls, obj: Any, **kwargs) -> None:
        """
        Set one or more properties.

        |lo_safe|

        Args:
            obj (Any): object to set properties for. Must support ``XPropertySet``
            **kwargs: Variable length Key value pairs used to set properties.

        Raises:
            MissingInterfaceError: if obj does not implement ``XPropertySet`` interface
            MultiError: If unable to set a property

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_SETTING` :eventref:`src-docs-props-event-setting`
                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_SET` :eventref:`src-docs-props-event-set`
                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_SET_ERROR` :eventref:`src-docs-props-event-set-error`

        Note:
            If a Event is canceled then that property is not set. No error occurs.

            If ``MultiError`` occurs only the properties that raised an error is part of the error object.
            The remaining properties will still be set.

            If ``KeyValCancelArgs.default`` is set to true then property is set to Default

        Note:
            If an error occurs a ``PROP_SET_ERROR`` is raised. If the event args are set to ``cancel`` or ``handled`` then the error is ignored.

        See Also:
            :py:meth:`~.props.Props.set_default`

        .. versionchanged:: 0.9.0
            Setting ``KeyValCancelArgs.default=False`` will set a property to its default.

        .. versionchanged:: 0.9.4
            Now event is raised if a property fails to set.
        """
        if not kwargs:
            return
        if mInfo.Info.is_type_interface(obj, "com.sun.star.beans.XPropertySet"):
            ps = cast(XPropertySet, obj)
        else:
            ps = mLo.Lo.qi(XPropertySet, obj, True)
        errs = []
        for key, value in kwargs.items():
            has_error = False
            cargs = KeyValCancelArgs(Props.set.__qualname__, key, value)
            cargs.event_data = obj
            _Events().trigger(PropsNamedEvent.PROP_SETTING, cargs)
            if cargs.cancel:
                continue
            try:
                if cargs.default:
                    cls.set_default(obj, cargs.key)
                elif cargs.key == "":
                    continue
                else:
                    ps.setPropertyValue(cargs.key, cargs.value)
            except (AttributeError, UnknownPropertyException) as e:
                # handle a LibreOffice bug
                try:
                    if not cls._set_by_attribute(obj, cargs.key, cargs.value):
                        raise e
                except Exception as ex:
                    has_error = True
                    if type(ex).__name__ == "com.sun.star.beans.UnknownPropertyException":
                        errs.append(mEx.PropertyNotFoundError(key, ex))
                    else:
                        errs.append(
                            mEx.UnKnownError(
                                f'Something went wrong. Could not find setPropertyValue attribute on property set. Tried setting "{key}" manually but failed'
                            )
                        )

            except PropertyVetoException as e:
                has_error = True
                errs.append(mEx.PropertySetError(f'Could not set readonly-property "{key}"', e))
            except Exception as e:
                has_error = True
                errs.append(Exception(f'Could not set property "{key}"', e))
            if has_error:
                error_args = KeyValCancelArgs.from_args(cargs)
                error_args.cancel = False
                error_args.handled = False
                _Events().trigger(PropsNamedEvent.PROP_SET_ERROR, error_args)
                if (error_args.handled or error_args.cancel) and errs:
                    _ = errs.pop()
            else:
                _Events().trigger(PropsNamedEvent.PROP_SET, KeyValArgs.from_args(cargs))  # type: ignore
        if errs:
            raise mEx.MultiError(errs)

    @classmethod
    def set_default(cls, obj: object, *prop_names: str) -> None:
        """
        Set one or more properties to default values.

        Args:
            obj (object): object to set properties for. Must support ``XPropertyState``
            **prop_names: Variable length of property names to set default.

        Raises:
            MissingInterfaceError: if obj does not implement ``XPropertyState`` interface
            MultiError: If unable to set a property

        Returns:
            None:

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_DEFAULT_SETTING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_DEFAULT_SET` :eventref:`src-docs-event`
                - :py:attr:`~.events.props_named_event.PropsNamedEvent.PROP_SET_DEFAULT_ERROR` :eventref:`src-docs-event-cancel`

        Note:
            Event data  is a dictionary of ``{"obj": obj, "name": name}`` with name being the property name currently being set.

            If a Event is canceled then that property is not set. No error occurs.

            If ``MultiError`` occurs only the properties that raised an error is part of the error object.
            The remaining properties will still be set to default.

        .. versionadded:: 0.9.0

        .. versionchanged:: 0.9.4
            Now event is raised if a property fails to set default.
        """
        ps = mLo.Lo.qi(XPropertyState, obj, True)
        errs = []
        for name in prop_names:
            has_error = False
            cargs = CancelEventArgs(Props.set.__qualname__)
            cargs.event_data = {"obj": obj, "name": name}
            _Events().trigger(PropsNamedEvent.PROP_DEFAULT_SETTING, cargs)
            if cargs.cancel:
                continue
            prop_name = cargs.event_data["name"]
            try:
                ps.setPropertyToDefault(prop_name)  # type: ignore
            except Exception as e:
                has_error = True
                errs.append(Exception(f'Could not set property Default "{prop_name}"', e))
            if has_error:
                error_args = CancelEventArgs.from_args(cargs)
                cargs.cancel = False
                cargs.handled = False
                _Events().trigger(PropsNamedEvent.PROP_SET_DEFAULT_ERROR, error_args)
                if (error_args.handled or error_args.cancel) and errs:
                    _ = errs.pop()
            else:
                _Events().trigger(PropsNamedEvent.PROP_DEFAULT_SET, EventArgs.from_args(cargs))
        if errs:
            raise mEx.MultiError(errs)

    # endregion ---------------- set properties -----------------------

    # region ------------------- get properties ------------------------
    @staticmethod
    def get_default(obj: object, prop_name: str, default: Any = gUtil.NULL_OBJ) -> Any:
        """
        Gets the default value of the property with the name Property Name.

        Args:
            obj (object): object to set properties for. Must support ``XPropertyState``
            prop_name (str): Property name to get default for.
            default (Any, optional): Return value if property value is ``None`` or ``obj`` does not support ``XPropertyState`` interface.

        Raises:
            MissingInterfaceError: If ``obj`` does not support ``XPropertyState`` interface and a ``default`` value is not provided.

        Returns:
            Any: If no default exists, is not known or is ``None``, then the return type is ``None``.
        """
        try:
            ps = mLo.Lo.qi(XPropertyState, obj, True)
            result = ps.getPropertyDefault(prop_name)
            return default if result is None and default is not gUtil.NULL_OBJ else result
        except mEx.MissingInterfaceError:
            if default is not gUtil.NULL_OBJ:
                return default
            raise

    # region get()
    @overload
    @classmethod
    def get(cls, obj: Any, name: str) -> Any:
        """
        Gets a property value from an object.

        |lo_safe|

        Args:
            obj (object): Object to get property from.
            name (str): Property Name.

        Returns:
            Any: Property value.
        """
        ...

    @overload
    @classmethod
    def get(cls, obj: Any, name: str, default: Any) -> Any:
        """
        Gets a property value from an object.

        |lo_safe|

        Args:
            obj (object): Object to get property from.
            name (str): Property Name.
            default (Any, optional): Return value if property value is ``None``.

        Returns:
            Any: Property value or default.
        """
        ...

    @classmethod
    def get(cls, obj: Any, name: str, default: Any = gUtil.NULL_OBJ) -> Any:
        """
        Gets a property value from an object.
        ``obj`` must support ``XPropertySet`` interface.

        |lo_safe|

        Args:
            obj (object): Object to get property from.
            name (str): Property Name.
            default (Any, optional): Return value if property value is ``None``.

        Raises:
            PropertyNotFoundError: If Property is not found and default was not set.
            PropertyError: If any other error occurs and default was not set.

        Returns:
            Any: Property value or default.

        Note:
            If a ``default`` is not set then an error is raised if property is not found.
            Otherwise, the ``default`` value is returned if property is not found.
        """
        try:
            if mInfo.Info.is_type_interface(obj, "com.sun.star.beans.XPropertySet"):
                ps = cast(XPropertySet, obj)
            else:
                ps = mLo.Lo.qi(XPropertySet, obj, True)
            try:
                val = ps.getPropertyValue(name)
            except (AttributeError, UnknownPropertyException) as e:
                # handle a LibreOffice bug
                success = False
                try:
                    success, val = cls._get_by_attribute(ps, name)
                except Exception as ex:
                    if type(e).__name__ == "com.sun.star.beans.UnknownPropertyException":
                        raise e from ex
                    raise mEx.UnKnownError(
                        f'Something went wrong. Could not find getPropertyValue attribute on property set. Tried getting "{name}" manually but failed.'
                    ) from ex

                if not success:
                    if type(e).__name__ == "com.sun.star.beans.UnknownPropertyException":
                        raise e
                    raise mEx.UnKnownError(
                        f'Something went wrong. Could not find getPropertyValue attribute on property set. Tried getting "{name}" manually but failed.'
                    ) from e
            # it is perfectly fine for a property to have a value of None
            return default if val is None and default is not gUtil.NULL_OBJ else val
        except UnknownPropertyException as exc:
            # the property name is not in the property set
            if default is not gUtil.NULL_OBJ:
                return default
            raise mEx.PropertyNotFoundError(name) from exc
        except mEx.UnKnownError:
            raise
        except Exception as exx:
            if default is not gUtil.NULL_OBJ:
                return default
            raise mEx.PropertyError(f'Error getting property: "{name}"') from exx

    # endregion get()

    # region    get_property()
    @overload
    @classmethod
    def get_property(cls, obj: object, name: str) -> object:
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
    @classmethod
    def get_property(cls, prop_set: XPropertySet, name: str) -> object:
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

    @classmethod
    def get_property(cls, *args, **kwargs) -> object:
        """
        Gets a property value from property set

        Args:
            obj (object): object that implements XPropertySet interface
            xprops (XPropertySet): property set
            name (str): property name

        Raises:
            PropertyNotFoundError: If Property is not found
            PropertyError: If any other error occurs.

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
            check = all(key in valid_keys for key in kwargs)
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
            raise TypeError("get_property() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        return cls.get(kargs[1], kargs[2])

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
        return tuple(props)  # type: ignore

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
        nms = [prop.Name for prop in props]
        return tuple(nms)

    @staticmethod
    def get_value(name: str, props: Iterable[PropertyValue]) -> object:
        """
        Get a property value from properties.

        |lo_safe|

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
                >>> from ooodev.loader.lo import Lo
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

        num_elements = in_acc.getCount()
        print(f"No. of elements: {num_elements}")
        if num_elements == 0:
            return
        for i in range(num_elements):
            try:
                # PropertyValue[] props = Lo.qi(PropertyValue[].class, inAcc.getByIndex(i));
                # above line is original java code.
                # perhaps there is a way to also include PropertyValue[] queryInterface
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
        Gets property values a a string.

        |lo_safe|

        Args:
            val (object): Values such as a iterable of iterable or
                object that implements XPropertySet or
                a string.

        Returns:
            str: A string representing properties.

        Example:
            .. code-block:: python

                >>> from ooodev.office.calc import Calc
                >>> from ooodev.utils.props import Props
                >>> from ooodev.loader.lo import Lo
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
                    lines.append(f"{p.Name} = {p.Value}")
                except AttributeError:
                    continue
            return "[\n    " + "\n    ".join(lines) + "\n  ]" if lines else "[]"

        def get_pv_str_f_arg(vals) -> str:
            # com.sun.star.sheet.FunctionArgument
            lines = []
            for p in vals:
                try:
                    if p.IsOptional:
                        lines.append(f"{p.Name} (optional)")
                    else:
                        lines.append(f"{p.Name}")
                except AttributeError:
                    continue
            return "[" + ", ".join(lines) + "]" if lines else "[]"

        def get_property_set_str(prop_set, props) -> str:
            lines = []
            for p in props:
                value = cls.get(prop_set, p.Name)
                lines.append(f"{p.Name} = {value}")
            return "[\n    " + "\n    ".join(lines) + "\n]"

        if val is None:
            return ""
        if isinstance(val, str):
            return val
        try:
            _ = iter(val)  # type: ignore
        except TypeError:
            # not iterable
            is_iter = False
        else:
            # iterable
            is_iter = True
        if is_iter:
            val_first = None
            with contextlib.suppress(Exception):
                val_first = val[0]  # type: ignore
                if isinstance(val_first, str):
                    # assume Iterable[str]
                    return ", ".join(val)  # type: ignore
            if val_first is None:
                return str(val)
            t_name = getattr(val_first, "typeName", None)
            if not t_name:
                return str(val)
            if t_name == "com.sun.star.beans.PropertyValue":
                return get_pv_str(val)
            if t_name == "com.sun.star.sheet.FunctionArgument":
                return get_pv_str_f_arg(val)
            return str(val)

            # assume Iterable[PropertyValue]
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
        Prints properties for an object to the console.

        |lo_safe|

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
        Prints properties to console.

        |lo_safe|

        Args:
            title (str): Title to use that is displayed in console
            props (Sequence[PropertyValue]): Properties to print
        """
        ...

    @overload
    @classmethod
    def show_props(cls, prop_kind: str, props_set: XPropertySet) -> None:
        """
        Prints properties to console.

        |lo_safe|

        Args:
            prop_kind (str): The kind of properties that is displayed in console
            props_set (XPropertySet): object the implements XPropertySet interface.
        """
        ...

    @classmethod
    def show_props(cls, *args, **kwargs) -> None:
        """
        Prints properties to console.

        |lo_safe|

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
            check = all(key in valid_keys for key in kwargs)
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
            raise TypeError("show_props() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mInfo.Info.is_type_interface(kargs[2], "com.sun.star.beans.XPropertySet"):
            # def show_props(prop_kind: str, props_set: XPropertySet)
            return cls._show_props_str_xproperty_set(prop_kind=kargs[1], props_set=kargs[2])
        else:
            # def show_props(title: str, props: Sequence[PropertyValue])
            return cls._show_props_str_props(title=kargs[1], props=kargs[2])

    @classmethod
    def _show_props_str_props(cls, title: str, props: Sequence[PropertyValue]) -> None:
        """LO Safe Method"""
        print(f'Properties for "{title}"":')
        if props is None:
            print("  none found")
            return
        for prop in props:
            print(f"  {prop.Name}: {cls.prop_value_to_string(prop.Value)}")
        print()

    @classmethod
    def _show_props_str_xproperty_set(cls, prop_kind: str, props_set: XPropertySet) -> None:
        """LO Safe Method"""
        props = cls.props_set_to_tuple(props_set)
        if props is None:
            print(f"No. {prop_kind} properties found")
            return
        print(f"{prop_kind} Properties")
        for prop in props:
            try:
                prop_value = props_set.getPropertyValue(prop.Name)
                print(f"  {prop.Name}: {prop_value}")
            except Exception as e:
                print(f'Unable to get property for "{prop.Name}"')
                print(f"  {e}")
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
        return xprops_info.getProperties()  # type: ignore

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
        return f"{p.Name}: {p.Type.typeName}"  # type: ignore

    # endregion ---------------- show properties of an Object ----------

    # region ------------------- others --------------------------------
    @classmethod
    def has(cls, obj: object, name: str) -> bool:
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
        try:
            return prop_set.getPropertySetInfo().hasPropertyByName(name)
        except AttributeError:
            # some object such as a chart data point seem to not properly implement XPropertySet
            # and does not have a getPropertySetInfo() method
            return hasattr(obj, name)

    has_property = has

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
        type_detect = mLo.Lo.create_instance_mcf(XTypeDetection, "com.sun.star.document.TypeDetection")
        if type_detect is None:
            print("No type detector reference")
            return

        name_access = mLo.Lo.qi(XNameAccess, type_detect, True)
        try:
            props = cast(Tuple[PropertyValue, ...], name_access.getByName(type))
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
