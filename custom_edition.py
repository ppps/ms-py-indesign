#!/usr/bin/env python3

import json
from pathlib import Path
import sys

masters_file = Path('masters.json')
pages_file = Path('pages.json')


def read_json(filename):
    with open(filename, encoding='utf-8') as f:
        loaded = json.load(f)
    return loaded


def read_masters(masters=None):
    if masters is None:
        masters = masters_file
    return read_json(masters)


def read_pages(pages=None):
    if pages is None:
        pages = pages_file
    return read_json(pages)


def main(custom_spec):
    spec_lines = [l.rstrip() for l in custom_spec.split('\n')
                  if not l.startswith('#')
                  if l.strip()]

    title = spec_lines[0]
    spec_lines = spec_lines[1:]

    masters = read_masters()
    pages_inventory = read_pages()

    reject = []
    accept = []

    for line in spec_lines:
        pn_str, master_name = line.split(maxsplit=1)

        try:
            page_number = int(pn_str)
        except ValueError:
            reject.append((line, 'First item must be an integer page number'))
            continue

        if master_name not in masters:
            reject.append((line, f'{master_name} is not in masters.json'))
            continue

        accept.append((line, page_number, master_name))

    print(f'# Found title: {title}')

    if accept:
        print('\n# These lines were interpreted OK:')
        for line, page, master in accept:
            print(line)

    if reject:
        print('\n# These lines were rejected:')
        for line, reason in reject:
            print(line)
            print(f'--> {reason}')
            print()
        print('# No special edition has been added to the generator.')
        print('# Please fix the above problems and retry.')
        return

    page_dicts = [{'master': master, 'page': page}
                  for _, page, master in accept]

    pages_inventory['Specials'][title] = page_dicts

    with open(pages_file, 'w', encoding='utf-8') as f:
        json.dump(pages_inventory, f, indent=2)

    print(f'\n# Added "{title}" to page generator Specials section')


if __name__ == '__main__':
    main(custom_spec=sys.stdin.read())
