from __future__ import annotations
from attr import attributes
import pytest
if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.meta.static_meta import StaticProperty, classinstanceproperty, classproperty
# some date util function rely on Lo.null_date
# Lo.null_date use a document to get actual null date.
# if Lo has no document then a null date of 1889/12/30 is used.


class MyClazz(metaclass=StaticProperty):
    # class variable are optional if class
    # properties are set in advanced.
    # eg:
    # MyClazz.ms_val = 10
    # MyClazz.c_val = 10
    _ms_val = 10
    _c_val = 20
    _read_val = 'read'
    @classproperty
    def ms_val(cls):
        return cls._ms_val

    @ms_val.setter
    def ms_val(cls, value):
        cls._cla_ms_valss = value
    
    @classinstanceproperty
    def c_val(cls):
        return cls._c_val

    @c_val.setter
    def c_val(cls, value):
        cls._c_val = value
    
    @classproperty
    def read_val(cls) -> str:
        # readonly class property, acts like const
        # will not raise an error
        return cls._read_val
    
    @classproperty
    def read_only_val(cls) -> str:
        try:
            return cls._read_only_val
        except AttributeError:
            cls._read_only_val = "read only"
        return cls._read_only_val

    @read_only_val.setter
    def read_only_val(cls, value) -> str:
        raise AttributeError("Attempt to modify read-only class property '%s'." % cls.name)

class TestClazz(metaclass=StaticProperty):
    @classproperty
    def class_(cls):
        return cls._class

    @class_.setter
    def class_(cls, value):
        cls._class = value

    @classinstanceproperty
    def class_instance(cls):
        return cls._class_instance

    @class_instance.setter
    def class_instance(cls, value):
        cls._class_instance = value

    @property
    def instance(self):
        return self._instance

    @instance.setter
    def instance(self, value):
        self._instance = value

def test_my_clazz() -> None:
    # classproperty are removed from instance properties
    # classinstanceproperty are not removed from instance properties
    # classinstanceproperty are readonly to instance
    assert MyClazz.ms_val == 10
    # MyClazz.c_val = 20
    assert MyClazz.c_val == 20
    mc = MyClazz()
    with pytest.raises(AttributeError):
        # classproperty are removed from instance properties
        assert mc.ms_val == 10
    mc.c_val == 33
    # instance can't be used after modifying class
    assert mc.c_val == 20

    MyClazz.c_val = 99
    assert MyClazz.c_val == 99
    assert mc.c_val == 99

    assert MyClazz.read_val == "read"
    MyClazz.read_val == 'yes'
    assert MyClazz.read_val == "read"
    assert MyClazz.read_only_val == 'read only'
    assert MyClazz.read_only_val == 'read only'
    with pytest.raises(AttributeError) as ex:
        MyClazz.read_only_val = "hello"

def test_clazz() -> None:
    tc = TestClazz()
    tc._instance = None
    tc.instance = True
    assert tc._instance is True
    assert tc.instance is True
    tc.instance = False
    assert tc._instance is False
    assert tc.instance is False
    
    TestClazz._instance = None
    TestClazz.instance = True
    TestClazz.instance = False
    assert TestClazz._instance is None
    tc._instance = True
    if TestClazz._instance is not True:
        print("instance can't be used after modifying class")
    
    TestClazz._class_instance = None
    TestClazz.class_instance = True
    assert TestClazz._class_instance is True
    TestClazz.class_instance = False
    assert TestClazz._class_instance is False
    
    tc = TestClazz()
    tc._class_instance = None
    tc.class_instance = True
    assert tc._class_instance is True
    assert TestClazz._class_instance is False
    tc.class_instance = False
    assert tc._class_instance is False
    
    TestClazz._class = None
    TestClazz.class_ = True
    assert TestClazz._class is True
    TestClazz.class_ = False
    assert TestClazz._class is False
    
    tc = TestClazz()
    tc._class = None
    tc.class_ = True
    assert tc._class is None
    assert TestClazz._class is False
    tc.class_ = False
    assert tc._class is None