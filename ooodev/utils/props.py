# coding: utf-8
# Python conversion of Props.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
"""make/get/set properties in an array"""
from __future__ import annotations
from typing import Iterable, Optional, Tuple, Union, TYPE_CHECKING
import uno

from com.sun.star.beans import Property, PropertyValue
from com.sun.star.beans import PropertyAttribute # type: ignore
from com.sun.star.ui import ItemType # type: ignore
from com.sun.star.ui import ItemStyle # type: ignore
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.uno import RuntimeException

# import module and not module content to avoid circular import issue.
# https://stackoverflow.com/questions/22187279/python-circular-importing
from . import lo as mLo

# Lo = m_lo.Lo

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertySet
    from com.sun.star.container import XIndexAccess
    from com.sun.star.lang import XServiceInfo
    from com.sun.star.document import XTypeDetection 
    from com.sun.star.container import XNameAccess
    from com.sun.star.beans import XMultiPropertySet

class Props:
    """make/get/set properties in an array"""
    @staticmethod
    def make_prop_value(name: Optional[str] = None, value: Optional[str] = None) -> PropertyValue:
        p: PropertyValue = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        if name is not None:
            p.Name = name
        if value is not None:
            p.Value = value
        return p
        
    
    @classmethod
    def make_bar_item(cls, cmd: str, item_name: str) -> Tuple[PropertyValue, PropertyValue, PropertyValue, PropertyValue, PropertyValue]:
        """
        propertiees for a toolbar item using a name and an image
        
        problem: image does not appear next to text on toolbar
        """
        p1 = cls.make_prop_value(name="CommandURL", value=cmd)
        p2 = cls.make_prop_value(name="Label", value=item_name)
        p3 = cls.make_prop_value(name="Type", value=ItemType.DEFAULT)
        p4 = cls.make_prop_value(name="Visible", value=True)
        p5 = cls.make_prop_value(name="Style")
        p5.Value = ItemStyle.DRAW_FLAT + ItemStyle.ALIGN_LEFT + ItemStyle.AUTO_SIZE + ItemStyle.ICON  + ItemStyle.TEXT
        
        return (p1, p2, p3, p4, p5)
    
    @classmethod
    def make_props(cls, **kwargs) -> Tuple[PropertyValue, ...]:
        """
        Make Properties

        KeyWord Args:
            kwargs (key, value): Each key, value pair is assigned to a PropertyValue.

        Returns:
            Tuple[PropertyValue, ...]: Tuple of Properties
        
        Notes:
            String properties such as ``Zoom.Value`` can be pass by constructing a dictionary
            and passing dictionary via exapnsion.
            
            Example Expansion

            .. code::
            
                p_dic = {
                    "Zoom.Value": 0,
                    "Zoom.ValueSet": 28703,
                    "Zoom.Type": 1
                    }
                props = Props.make_props(**p_dic)
        """
        lst = []
        for k, v in kwargs.items():
            lst.append(
                cls.make_prop_value(name=k, value=v)
            )
        return tuple(lst)

    @staticmethod
    def set_prop(props: Iterable[PropertyValue], name: str, value: object) -> None:
        if props is None:
            print(f"Property array is null; cannot set {name}")
            return
        for prop in props:
            if prop.Name == name:
                prop.Value = value
                return
        print("{name} not found")
    
    @staticmethod
    def get_prop(props: Iterable[PropertyValue], name: str) -> Union[object, None]:
        if props is None:
            print(f"Property array is null; cannot get {name}")
            return None
        for prop in props:
            if prop.Name == name:
                return prop.Value
        print(f"{name} not found")
        return None
    
    @staticmethod
    def set_property(prop_set: XPropertySet, name: str, value: object) -> None:
        if prop_set is None:
            print(f"Property set is null; cannot set '{name}'")
            return
        try:
            prop_set.setPropertyValue(name, value)
        except IllegalArgumentException as e:
            print(f"Property '{name}' argument is illegal")
        except Exception as e:
            print(f"Coul not set property '{name}': {e}")
    
    @classmethod
    def set_properties(cls, prop_set: XPropertySet, from_props: XPropertySet) -> None:
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
                print(f"Could not set property '{itm}': {e}")
    
    @staticmethod
    def get_property(x_props: XPropertySet, name: str) -> Union[object, None]:
        value = None
        try:
            value = x_props.getPropertyValue(name)
        except RuntimeException as e:
            print(f"Could not get runtime property '{name}': {e}")
        except Exception as e:
            print(f"Could not get property '{name}': {e}")
        return value
    
    @staticmethod
    def get_properties(prop_set: XPropertySet) -> Tuple[PropertyValue, ...]:
        props = list(prop_set.getPropertySetInfo().getProperties())
        props.sort(key=lambda prop: prop.Name)
        return tuple(props)
    
    @staticmethod
    def get_prop_names(prop_set: XPropertySet) -> Tuple[str, ...]:
        props = prop_set.getPropertySetInfo().getProperties()
        nms = []
        for prop in props:
            nms.append(prop.Name)
        return tuple(nms)
    
    @staticmethod
    def has_property(prop_set: XPropertySet, name: str) -> bool:
        return prop_set.getPropertySetInfo().hasPropertyByName(name)
    
    @staticmethod
    def get_value(name: str, props: Iterable[PropertyValue]) -> Union[object, None]:
        for prop in props:
            if prop.Name == name:
                return prop.Value
        print(f"{name} not found")
        return None
    
    # ------------------ show properties array ------------------------------
    
    @classmethod
    def show_indexed_props(cls, title: str, obj: XIndexAccess) -> None:
        if not hasattr(obj, "getCount"):
            print("Could not convert object to an IndexAccess container")
            return
        num_elems = obj.getCount()
        print(f"No. of elements: {num_elems}")
        if num_elems == 0:
            return
        for i in range(num_elems -1):
            try:
                props: Tuple[PropertyValue, ...] = obj.getByIndex(i)
                cls.show_props(f"Elem {i}", props)
                print("----")
            except Exception as e:
                print(f"Could not get elem {i}: {e}")

    @classmethod
    def show_props(cls, title: str, props: Iterable[PropertyValue]) -> None:
        print(f"Properties for '{title}':")
        if props is None:
            print("  none found")
            return
        for prop in props:
            print(f"  {prop.Name}: {cls.prop_value_to_string(prop.Value)}")
        print()
    
    @staticmethod
    def prop_value_to_string(val: Union[Iterable[PropertyValue], Iterable[str], object]) -> str:
        if val is None:
            return ''
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
            
            if len(val) == 0:
                return ''
            if isinstance(val[0], str):
                # assume Iterable[str]
                return ', '.join(val)
            # assume Iterable[PropertyValue]
            s = "["
            for p in val:
                s += f"\n    {p.Name} = {p.Value}"
            s += "\n  ]"
            return s
        else:
            return str(val)

    @classmethod
    def show_values(cls, name: str, props: Iterable[PropertyValue]) -> None:
        for prop in props:
            if prop.Name == name:
                print(f"{prop.Name}: {cls.prop_value_to_string(prop.Value)}")
                return
        print(f"{name} not found")

    # ------------------- show properties of an Object ----------------
    
    @classmethod
    def show_obj_props(cls, prop_kind:str, obj: XServiceInfo) -> None:
        try:
            if obj.supportsService('com.sun.star.beans.XPropertySet'):
                cls.show_props(prop_kind, obj)
                return
        except AttributeError:
            pass
        
        print(f"No properties found for {prop_kind}")
    
    @classmethod
    def show_props(cls, prop_kind: str, props_set: XPropertySet) -> None:
        props = cls.props_set_to_array(props_set)
        if props is None:
            print(f"No {prop_kind} properties found")
            return
        lst = list(props)
        lst.sort(key=lambda prop: prop.Name)
        print(f"{prop_kind} Properties")
        for prop in lst:
            prop_value = cls.get_property(props_set, prop.Name)
            print(f"  {prop.Name} == {prop_value}")
        print()
        
    @staticmethod
    def props_set_to_array(xprops: XPropertySet) -> Tuple[Property, ...]:
        if xprops is None:
            return ()
        xprops_info = xprops.getPropertySetInfo()
        return xprops_info.getProperties()
    
    @staticmethod
    def show_property(p: Property) -> str:
        return f"{p.Name}: {p.Type.getTypeName()}"
    
    # --------------------------- others ------------------------------
    
    @classmethod
    def show_doc_type_props(cls, type: str) -> None:
        if type is None:
            print('type is None')
            return
        xtype_detect: Union[XTypeDetection, XNameAccess] = mLo.Lo.create_instance_mcf("com.sun.star.document.TypeDetection")
        if xtype_detect is None:
            print('No type detector reference')
            return
        try:
            props = xtype_detect.getByName(type)
            cls.show_props(type, props)
        except Exception:
            print(f"No properties for '{type}'")
    
    @staticmethod
    def get_bound_props(props_set: XMultiPropertySet) -> Tuple[str, ...]:
        props = props_set.getPropertySetInfo().getProperties()
        names = []
        for p in props:
            is_writable = ((p.Attributes & PropertyAttribute.READONLY) == 0)
            is_not_null = ((p.Attributes & PropertyAttribute.MAYBEVOID) == 0)
            is_bound = ((p.Attributes & PropertyAttribute.BOUND) == 0)
            if is_writable and is_not_null and is_bound:
                names.append(p.Name)
        names_len = len(names)
        if names_len == 0:
            print("No suitable properties were found")
            return ()
        print(f"No. of suitable properties: {names_len}")
        return tuple(names)