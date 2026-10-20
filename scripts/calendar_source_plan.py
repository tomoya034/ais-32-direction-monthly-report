"""Calendar rows and actual source identities are different contracts."""
import calendar
from datetime import date


def source_calendar_plan(year, month, logical_days):
    if logical_days is None:
        raise ValueError("Complete source logical-day plan required")
    source_days = [date.fromisoformat(value) for value in logical_days]
    if source_days != sorted(set(source_days)):
        raise ValueError("Source logical-day plan must be unique and chronological")
    if any((day.year,day.month)!=(year,month) for day in source_days):
        raise ValueError("Source logical-day plan contains another month")
    calendar_days = [date(year,month,day) for day in range(1,calendar.monthrange(year,month)[1]+1)]
    return source_days,calendar_days
