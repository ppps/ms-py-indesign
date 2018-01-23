#!/usr/bin/env python3
"""
InDesign Page Generator

Usage:
    gen.py --desk=DESK --date=DATE
"""

ID_VERSION = 'CS5.5'

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
    script = f'''\
tell application "Adobe InDesign {ID_VERSION}"
  tell the front document
    set the contents of every text frame whose label is "{frame_name}" to "{text}"
  end tell
end tell
'''
    run_applescript(script)


if __name__ == '__main__':
    set_frame_contents('X', 'contents')
