#!/usr/bin/env python3

from datetime import datetime

import gen


def test_format_page_date():
    """Check date formatting with format_page_date

    It should return a format matching:
        %A\n%B %-d %Y
    For weekdays and Sundays.

    That is, the weekday name in full on one line, and the month,
    date and year on the following line:

    Wednesday
    January 1 1930

    Saturdays are tested separately.
    """
    cases = [
        datetime(2018, 1, 23),
        datetime(2017, 11, 6),
        datetime(2016, 9, 25)]
    for case in cases:
        assert gen.format_page_date(case) == case.strftime('%A\n%B %-d %Y')


def test_format_page_date_weekend():
    """Check weekend formatting with format_page_date

    Saturday editions that span the weekend also include Sunday’s date
    in the string.

    The simplest case is just two day names and two dates:
        Saturday/Sunday
        January 27-28 2018

    But it should also handle month boundaries:
        Saturday/Sunday
        March 31-April 1 2018

    And year boundaries:
        Saturday/Sunday
        December 31 2016-January 1 2017
    """
    cases = [
        (datetime(2018, 1, 27), 'Saturday/Sunday\nJanuary 27-28 2018'),
        (datetime(2018, 3, 31), 'Saturday/Sunday\nMarch 31-April 1 2018'),
        (datetime(2016, 12, 31),
         'Saturday/Sunday\nDecember 31-January 1 2016-2017')
        ]
    for case, expected in cases:
        assert gen.format_page_date(case) == expected


def test_file_date_formatting():
    """format_file_date should return a six-digit DDMMYY string"""
    cases = [
        datetime(2018, 1, 27),
        datetime(2018, 3, 31),
        datetime(2016, 12, 31)]
    for case in cases:
        assert gen.format_file_date(case) == case.strftime('%d%m%y')
