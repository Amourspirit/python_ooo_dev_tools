# coding: utf-8
from __future__ import annotations
import datetime
import numbers
from typing import cast
from . import lo as mLo

class DateUtil:
     # region --------------- convert methods ---------------------------
    @staticmethod
    def date_from_number(value: numbers.Number) -> datetime.datetime | None:
        """
        Converts a float value to corresponding datetime instance.

        Args:
            value (Number): number to convert to date and time

        Returns:
            datetime | None: Date time instance on success; Otherwise, None
        """
        if not isinstance(value, numbers.Real):
            return None
        delta = datetime.timedelta(days=value)
        dnull = cast(datetime.datetime, mLo.Lo.null_date)
        return dnull + delta
    
    @staticmethod
    def date_to_number(date: datetime.datetime | datetime.date) -> float | None:
        """
        Converts a date or datetime instance to a corresponding float value.

        Args:
            date (datetime | date): date or date time to convert to float

        Returns:
            float | None: date as float on success; Otherwise, None
        """
        dnull = cast(datetime.datetime, mLo.Lo.null_date)
        if isinstance(date, datetime.datetime):
            delta = date - dnull
        elif isinstance(date, datetime.date):
            delta = date - dnull.date()
        else:
            print("date is incorrect type")
            # raise TypeError(date)
            return None
        return delta.days + delta.seconds / (24.0 * 60 * 60)
    
    @staticmethod
    def time_from_number(value: numbers.Number) -> datetime.time | None:
        """
        Converts a float value to corresponding time instance.
        
        Only Hours, Minutes and seconds are are used in conversion.

        Args:
            value (Number): Number such a float or int to convert

        Returns:
            time | None: Value as time on success; Otherwise, None
        """
        if not isinstance(value, numbers.Real):
            return None
        delta = datetime.timedelta(days=value)
        minutes, second = divmod(delta.seconds, 60)
        hour, minute = divmod(minutes, 60)
        return datetime.time(hour, minute, second)
    
    @staticmethod
    def time_to_number(time: datetime.time) -> float | None:
        """
        Converts a time instance to a corresponding float value.
        
        Only Hours, Minutes and seconds are are used in conversion.

        Args:
            time (datetime.time): time to convert

        Returns:
            float | None: time as float on success; Otherwise, None
        """
        if not isinstance(time, datetime.time):
            print("time is incorrect type")
            # raise TypeError(time)
            return None
        return ((time.second / 60.0 + time.minute) / 60.0 + time.hour) / 24.0

    @staticmethod
    def date_time_str(dt: datetime.datetime) -> str:
        """
        returns a formated date and time as string

        Args:
            dt (datetime): date time

        Returns:
            str: formatted date string such as 'Jun 05, 2022 20:15'
        """
        return dt.strftime("%b %d, %Y %H:%M")
    # endregion ------------ convert methods ---------------------------