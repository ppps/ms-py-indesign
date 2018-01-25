#!/usr/bin/env python3
"""
InDesign Page Generator

Usage:
    gen.py --master=MASTER --pages_dir=DIR
"""

import copy
from datetime import datetime, timedelta
import json
import logging
from pathlib import Path
import re
import subprocess
import sys

APP_DIR = Path(__file__).parent

from docopt import docopt

logging.basicConfig(
    format='%(asctime)s  %(levelname)-10s %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    level=logging.DEBUG)
log = logging.getLogger(__name__)


def remove_zero_padded_dates(date_string):
    """Mimic %-d on unsupported platforms by trimming zero-padding

    For example:
        January 01 2018
    Becomes:
        January 1 2018
    """
    return re.sub(r' 0(\d)', r' \1', date_string)


def run_applescript(script_str):
    """Encode and run the AppleScript in script_str

    An AppleScript execution error will cause the program to
    quit immediately. This is logged as a fatal error.
    """
    osa = subprocess.Popen(['osascript', '-'],
                           stdin=subprocess.PIPE,
                           stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE)
    result = osa.communicate(script_str.encode('utf-8'))

    decoded = [stream.decode('utf-8').rstrip() for stream in result]
    stdout, stderr = decoded
    if any(decoded):
        log.debug('AppleScript output: %s', decoded)
    if 'execution error' in stderr:
        # Program cannot continue running as it is in an unknown state
        log.critical('AppleScript: ' + stderr.split('execution error: ')[-1])
        sys.exit()

    return stdout if stdout else stderr


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
    script = ('set the contents of every text frame whose label is '
              f'"{frame_name}" to "{text}"')
    wrap_and_run(script)


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
        return ('Saturday/Sunday\n'
                f'December 31-January 1 {saturday.year}-{sunday.year}')
    elif saturday.month != sunday.month:
        date = f'Saturday/Sunday\n{saturday:%B %d}-{sunday:%B %d %Y}'
    else:
        date = f'Saturday/Sunday\n{saturday:%B %d}-{sunday:%d %Y}'
    return remove_zero_padded_dates(date)


def format_page_date(edition_date, sat_spans_weekend=True):
    """Return a string to represent the date on the page

    The format used is:
        Tuesday
        January 23 2018
    Or:
        %A\n%B %-d %Y

    If sat_spans_weekend, Saturday dates return a joint Sat/Sun string:
        Saturday/Sunday
        January 27-28 2018

    This also handles month and year boundaries. (A separate
    format_weekend_date function is used to deal with the complexity.)
    """
    if sat_spans_weekend and edition_date.isoweekday() == 6:
        return _format_page_date_for_weekend(edition_date)
    date = edition_date.strftime('%A\n%B %d %Y')
    return remove_zero_padded_dates(date)


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
            'make new spread with properties {applied master:master spread "' +
            master_name + '"}')
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
set locked of layer "Furniture" to true
set active layer to "Work"
save to POSIX file "{path}"
'''
    wrap_and_run(script)


def format_file_path(edition_date, page_number, slug,
                     spread: bool, pages_root):
    """Work out where to save the new InDesign file"""
    date_string = edition_date.strftime('%Y-%m-%d %A %b %d')
    date_string = remove_zero_padded_dates(date_string)
    edition_directory = pages_root.joinpath(date_string)
    edition_directory.mkdir(parents=True, exist_ok=True)

    if spread:
        str_num = '-'.join(map(str, [page_number, page_number + 1]))
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


def override_master_items(master_name, spread=False):
    """Override items from the work layer on the master"""
    script = 'try\n'
    page_nums = [2, 3] if spread else [1]
    for num in page_nums:
        script += (
            f'override (every item of master page items of page {num}'
            'whose item layer\'s name is "Work") destination page '
            f'page {num}\n')
    script += 'end try'
    return wrap_and_run(script)


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
    override_master_items(master_name, spread=spread)
    save_location = format_file_path(edition_date, page_number, slug, spread,
                                     pages_root)
    save_file(path=save_location)
    close_active_document()


def load_masters_json(masters_file='masters.json'):
    """Load a JSON file containing the specification for the master pages"""
    with open(APP_DIR.joinpath(masters_file)) as json_file:
        masters = json.load(json_file)
    return masters


def load_generators_json(pages_file='pages.json'):
    """Load a JSON file showing the page sets available to be generated"""
    with open(APP_DIR.joinpath(pages_file)) as json_file:
        pages = json.load(json_file)
    return pages


def construct_page_specifications(pages_dict, masters_dict):
    """Construct dict by adding details from masters to page instructions

    The pages dict only contains the name of the master to use and the
    page number, which is not enough to generate the page. This function
    fills in the needed details from the masters dict.

    This allows the file containing the page-generating instructions
    to be kept reasonably clear, allowing for easier editing, and
    removes repetition.
    """
    detailed_dict = copy.deepcopy(pages_dict)
    for desk in detailed_dict:
        for page_list in detailed_dict[desk].values():
            for page in page_list:
                master = masters[page['master']]
                page['slug'] = master['slug']
                page['spread'] = master['spread']
    return detailed_dict


def wrap_seq_for_applescript(seq):
    """Wrap a Python sequence in braces and quotes for use in AppleScript"""
    quoted = [f'"{item}"' for item in seq]
    joined = ', '.join(quoted)
    wrapped = '{' + joined + '}'
    return wrapped


def prompt_for_list_selection(sequence, multiple_selections=False):
    """Wrap the items of sequence and ask the user to choose from them

    This always returns a list, even for a single selection.

    If the user chooses cancel this function will exist the program
    using sys.exit

    Raises ValueError if multiple_selections is True and any of the
    items of sequence contain ', ' — as this is what AppleScript uses
    to separate multiple (unquoted) values in the string returned.
    """
    if (multiple_selections
            and any(', ' in str(s) for s in sequence)):
        raise ValueError(
            '", " is disallowed in sequence items when multiple'
            ' selection is enabled.')
    script = f'''\
tell application "System Events"
  choose from list {wrap_seq_for_applescript(sequence)}{' with multiple selections allowed' if multiple_selections else ''}
end tell
'''
    result = run_applescript(script)
    if result == 'false':
        log.debug('User cancelled list selection')
        sys.exit()
    if multiple_selections:
        # Split the multiple values returned from AppleScript
        return result.split(', ')
    else:
        return [result]


def prompt_for_text_input(message, default=''):
    """Prompt the user to input text, or confirm the default

    If the user cancels the dialog this function will exit the program
    using sys.exit
    """
    result = run_applescript(f'''\
tell application "System Events"
  display dialog "{message}" default answer "{default}"
end tell
''')
    if 'execution error: User canceled. (-128)' in result:
        log.debug('User cancelled text input')
        sys.exit()
    return result.split('text returned:')[-1]


def prompt_for_date(offset=1):
    """Prompt the user to enter the date and return a datetime object

    By default the date shown is tomorrow's date (with an offset of one
    day, controlled by the offset parameter).
    """
    tomorrow_iso = (datetime.today() + timedelta(1)).strftime('%Y-%m-%d')
    result = prompt_for_text_input(
        'Enter the date in ISO format (YYYY-MM-DD).\n' +
        'The default is tomorrow',
        default=tomorrow_iso)
    date_match = re.search(r'(\d{4}-\d{2}-\d{2})', result)
    if not date_match:
        log.critical('Could not find acceptable date')
        sys.exit()
    return datetime.strptime(date_match.group(1), '%Y-%m-%d')


if __name__ == '__main__':
    args = docopt(__doc__)

    pages_root = Path(args['--pages_dir']).expanduser().resolve()
    master_file = Path(args['--master']).expanduser().resolve()

    masters = load_masters_json()
    pages = load_generators_json()
    pages = construct_page_specifications(pages, masters)

    desk = prompt_for_list_selection(pages)[0]
    date = prompt_for_date()

    try:
        to_generate = prompt_for_list_selection(
            pages[desk], multiple_selections=True)
    except ValueError:
        log.critical('Malformed page set name. Cannot continue.')
        sys.exit()

    for page_set_name in to_generate:
        for page in pages[desk][page_set_name]:
            create_from_master(
                master_name=page['master'],
                spread=page['spread'],
                slug=page['slug'],
                page_number=page['page'],
                edition_date=date,
                master_file=master_file,
                pages_root=pages_root)
