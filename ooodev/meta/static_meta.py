# coding: utf-8
class classproperty(property):
    pass

class classinstanceproperty(property):
    pass

class StaticProperty(type):
    def __new__(self, name, bases, props):
        class_properties = {}
        to_remove = {}
        for key, value in props.items():
            if isinstance(value, (classproperty, classinstanceproperty)):
                class_properties[key] = value
                if isinstance(value, classproperty):
                    to_remove[key] = value

        for key in to_remove:
            props.pop(key)

        HoistMeta = type('HoistMeta', (type,), class_properties)
        return HoistMeta(name, bases, props)