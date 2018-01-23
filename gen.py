#!/usr/bin/env python3
"""
InDesign Page Generator

Usage:
    gen.py --desk=DESK --date=DATE
"""

from datetime import datetime, timedelta
import json
import subprocess

from docopt import docopt


def run_applescript(script_str):
    """Encode and run the AppleScript in script_str"""
    osa = subprocess.Popen(['osascript', '-'],
                           stdin=subprocess.PIPE,
                           stdout=subprocess.PIPE,
                           stderr=subprocess.DEVNULL)
    result = osa.communicate(script_str.encode('utf-8'))[0].decode('utf-8')
    return result.rstrip()


def set_frame_contents(frame_name, text):
    """Set the contents of text frames in the active InDesign document

    frame_name corresponds to the script label of the frame in InDesign.

    All frames with the same label have the contents set to `text`.
    """
    script = f'''\
tell application "Adobe InDesign CS4"
  tell the front document
    set the contents of text frame "{frame_name}" to "{text}"
  end tell
end tell
'''
    run_applescript(script)


def _format_page_date_for_weekend(edition_date):
    """Format two-day weekend edition dates

    Saturday editions that include Sunday require special treatment.

    The simplest case is just two day names and two dates:
        Saturday/Sunday January 27-28 2018

    But dates can span month boundaries:
        Saturday/Sunday March 31-April 1 2018

    And year boundaries:
        Saturday/Sunday December 31 2016-January 1 2017
    """
    saturday = edition_date
    sunday = edition_date + timedelta(1)
    if saturday.year != sunday.year:
        return ('Saturday/Sunday December 31-January 1 '
                f'{saturday.year}-{sunday.year}')
    elif saturday.month != sunday.month:
        # Just %d for Saturday because it’s date is never less than 10
        return f'Saturday/Sunday {saturday:%B %d}-{sunday:%B %-d %Y}'
    else:
        return f'Saturday/Sunday {saturday:%B %-d}-{sunday:%-d %Y}'


def format_page_date(edition_date, sat_spans_weekend=True):
    """Return a string to represent the date on the page

    The format used is:
        Tuesday January 23 2018
    Or:
        %A %B %-d %Y

    If sat_spans_weekend, Saturday dates return a joint Sat/Sun string:
        Saturday/Sunday January 27-28 2018

    This also handles month and year boundaries. (A separate
    format_weekend_date function is used to deal with the complexity.)
    """
    if sat_spans_weekend and edition_date.isoweekday() == 6:
        return _format_page_date_for_weekend(edition_date)
    return edition_date.strftime('%A %B %-d %Y')


def format_file_date(edition_date):
    """Return a string DDMMYY for use in the filename"""
    return edition_date.strftime('%d%m%y')


if __name__ == '__main__':
    set_frame_contents('X', 'contents')
