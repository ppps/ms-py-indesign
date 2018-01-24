#!/usr/bin/env python3
"""
InDesign Page Generator

Usage:
    gen.py --master=MASTER --pages_dir=DIR
"""

from datetime import datetime, timedelta
import json
import logging
from pathlib import Path
import subprocess

from docopt import docopt

logging.basicConfig(
    format='%(asctime)s  %(levelname)-10s %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    level=logging.DEBUG)
log = logging.getLogger(__name__)


# SERVER_PAGES_DIR = Path('/Volumes/Server/Pages/')
SERVER_PAGES_DIR = Path('/Users/robjwells/Desktop/Pages/')
SERVER_MASTER = Path('/Volumes/Server/Production resources/Master pages/'
                     'CS4 Master.indd')
SERVER_MASTER = Path('/Users/robjwells/Desktop/CS4 Master.indd')


def run_applescript(script_str):
    """Encode and run the AppleScript in script_str"""
    osa = subprocess.Popen(['osascript', '-'],
                           stdin=subprocess.PIPE,
                           stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE)
    result = osa.communicate(script_str.encode('utf-8'))

    decoded = [stream.decode('utf-8').rstrip() for stream in result]
    if any(decoded):
        log.debug(decoded)

    return decoded[0]


def wrap_and_run(script):
    """Wrap the InDesign script in appropriate tell blocks and run it

    The script is wrapped in the boilerplate block for addressing
    the active application in Adobe InDesign. This is to cut down
    on repetition in this file.

    The result of the AppleScript runner is returned
    """
    return run_applescript(f'''
tell application "Adobe InDesign CS4"
  tell the active document
    {script}
  end tell
end tell
''')


def set_frame_contents(frame_name, text):
    """Set the contents of text frames in the active InDesign document

    frame_name corresponds to the script label of the frame in InDesign.

    All frames with the same label have the contents set to `text`.
    """
    wrap_and_run(
        'set the contents of every text frame whose label is '
        f'"{frame_name}" to "{text}"')


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
        # Just %d for Saturday because its date is never less than 10
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


def apply_master(master_name: str, spread: bool):
    """Create a working page from the specified master page

    If `spread` is True then a new spread is created in the document,
    underneath the existing single page (so that the working pages
    are 2 & 3 in the document.

    This function applies the master page, then overrides items that
    need to be set later.
    """
    if not spread:
        apply_master_script = (
            f'set applied master of page 1 to master spread "{master_name}"')
    else:
        apply_master_script = (
            'make new spread with properties {applied master:master spread ' +
            master_name + '}')
    wrap_and_run(apply_master_script)


def set_date_on_page(edition_date):
    """Set the content of the current document’s `Edition date` frames

    edition_date should be a datetime object (or something with strftime)
    """
    set_frame_contents('Edition date', format_page_date(edition_date))


def set_spread_page_numbers(left_page_number):
    """Set the page numbers on both halves of a spread"""
    set_frame_contents('L-Page number', left_page_number)
    set_frame_contents('R-Page number', left_page_number + 1)


def set_single_page_number(page_number):
    """Set the page numbers on a single page"""
    set_frame_contents('Page number', page_number)


def save_file(path):
    """Save the active document to the provided path

    path should be a pathlib.Path object (as the path
    needs to be resolved, and .resolve() is called on it.)
    """
    path = path.resolve()
    script = f'''\
set locked of layer "Date and page number" to true
set locked of layer "Furniture" to true
set active layer to "Work"
save to POSIX file "{path}"
'''
    wrap_and_run(script)


def format_file_path(edition_date, page_number, slug,
                     spread: bool, pages_root):
    """Work out where to save the new InDesign file"""
    edition_directory = pages_root.joinpath(
        edition_date.strftime('%Y-%m-%d %A %b %-d'))
    edition_directory.mkdir(exist_ok=True)

    if spread:
        str_num = '-'.join([page_number, page_number + 1])
    else:
        str_num = str(page_number)

    file_date = format_file_date(edition_date)

    return edition_directory.joinpath(f'{str_num}_{slug}_{file_date}.indd')


def open_master(master_file):
    """Open the master InDesign file for page creation"""
    run_applescript(f'''
tell application "Adobe InDesign CS4"
  open POSIX file "{master_file}"
end tell
''')


def close_active_document():
    """Close and save the current active InDesign document"""
    wrap_and_run('close saving yes')


def create_from_master(master_name: str, spread: bool, slug,
                       edition_date: datetime, page_number: int,
                       master_file, pages_root):
    """Create a new working document from a master page"""
    open_master(master_file)
    apply_master(master_name, spread)
    set_date_on_page(edition_date)
    if spread:
        set_spread_page_numbers(page_number)
    else:
        set_single_page_number(page_number)
    save_location = format_file_path(edition_date, page_number, slug, spread,
                                     pages_root)
    save_file(path=save_location)
    close_active_document()


def load_masters_json(masters_file='masters.json'):
    """Load a JSON file containing the specification for the master pages"""
    with open(masters_file) as json_file:
        masters = json.load(json_file)
    return masters


def load_generators_json(pages_file='pages.json'):
    """Load a JSON file showing the page sets available to be generated"""
    with open(pages_file) as json_file:
        pages = json.load(json_file)
    return pages


if __name__ == '__main__':
    args = docopt(__doc__)
    log.debug(args)

    pages_root = Path(args['--pages_dir']).expanduser().resolve()
    master_file = Path(args['--master']).expanduser().resolve()

#     print(load_masters_json())
#     print(load_generators_json())

#     create_from_master(
#         master_name='Feat-Letters-L',
#         spread=False,
#         slug='Letters',
#         edition_date=datetime.today(),
#         page_number=14,
#         master_file=master_file,
#         pages_root=pages_root)
