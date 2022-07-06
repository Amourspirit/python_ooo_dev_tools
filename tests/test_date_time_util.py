from __future__ import annotations
import pytest
if __name__ == "__main__":
    pytest.main([__file__])
import datetime
from ooodev.utils.lo import Lo
from ooodev.utils.date_time_util import DateUtil

# some date util function rely on Lo.null_date
# Lo.null_date use a document to get actual null date.
# if Lo has no document then a null date of 1889/12/30 is used.

def test_using_calc(loader) -> None:
    from ooodev.office.calc import Calc
    doc = Calc.create_doc(loader=loader)
    assert doc is not None, "Could not create new document"
    d = datetime.datetime(year=2022, month=6,day=1,hour=9, tzinfo=datetime.timezone.utc)
    num = DateUtil.date_to_number(d)
    # because null date is not fixes the result may vary
    assert num > 0.0
    c_date = DateUtil.date_from_number(value=num)
    assert c_date is not None
    assert c_date == d
    
    t = datetime.time(hour=11, minute=11, second=11, tzinfo=datetime.timezone.utc)
    t_num = DateUtil.time_to_number(time=t)
    assert t_num == pytest.approx(0.46609953703703705, rel=1e-6)
    c_time = DateUtil.time_from_number(t_num)
    assert t == c_time
    Lo.close(closeable=doc, deliver_ownership=False)
    
def test_using_writer(loader) -> None:
    from ooodev.office.write import Write
    doc = Write.create_doc(loader=loader)
    assert doc is not None, "Could not create new document"
    d = datetime.datetime(year=2022, month=6,day=1,hour=9, tzinfo=datetime.timezone.utc)
    num = DateUtil.date_to_number(d)
    # because null date is not fixes the result may vary
    assert num > 0.0
    c_date = DateUtil.date_from_number(value=num)
    assert c_date is not None
    assert c_date == d
    
    t = datetime.time(hour=11, minute=11, second=11, tzinfo=datetime.timezone.utc)
    t_num = DateUtil.time_to_number(time=t)
    assert t_num == pytest.approx(0.46609953703703705, rel=1e-6)
    c_time = DateUtil.time_from_number(t_num)
    assert t == c_time
    Lo.close(closeable=doc, deliver_ownership=False)

def test_using_no_doc() -> None:
    d = datetime.datetime(year=2022, month=6,day=1,hour=9, tzinfo=datetime.timezone.utc)
    num = DateUtil.date_to_number(d)
    assert num > 0
    c_date = DateUtil.date_from_number(value=num)
    assert c_date is not None
    assert c_date == d
    
    t = datetime.time(hour=11, minute=11, second=11, tzinfo=datetime.timezone.utc)
    t_num = DateUtil.time_to_number(time=t)
    assert t_num == pytest.approx(0.46609953703703705, rel=1e-6)
    c_time = DateUtil.time_from_number(t_num)
    assert t == c_time


