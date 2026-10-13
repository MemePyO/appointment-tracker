"""Auto-fill L days: every Saturday and Sunday of a given month/year.

Pure logic, no UI — shared by the desktop app and (later) the web bridge.
"""

from calendar import monthrange
from datetime import date

import lang
from workbook import MONTHS


def _check_year(year):
    if year is None or not str(year).strip():
        return date.today().year
    try:
        year = int(str(year).strip())
    except ValueError:
        raise ValueError(lang.t("err_year_whole"))
    if not 2000 <= year <= 2100:
        raise ValueError(lang.t("err_year_range"))
    return year


def _month_number(month):
    text = str(month or "").strip().casefold()
    if not text:
        raise ValueError(lang.t("err_month_empty"))
    if text.isdigit():
        num = int(text)
        if not 1 <= num <= 12:
            raise ValueError(lang.t("err_month_number", text=text))
        return num
    num = MONTHS.get(text)
    if num is None:
        raise ValueError(lang.t("err_month_unknown", month=month))
    return num


def weekend_days(month, year=None):
    """Return sorted day numbers of every Saturday and Sunday in the month."""
    num = _month_number(month)
    year = _check_year(year)
    n_days = monthrange(year, num)[1]
    days = []
    for day in range(1, n_days + 1):
        if date(year, num, day).weekday() >= 5:  # Mon=0 … Sat=5, Sun=6
            days.append(day)
    return days
