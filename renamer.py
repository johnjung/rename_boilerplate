'''
This program is free software: you can redistribute it and/or modify it under
the terms of the GNU General Public License as published by the Free Software
Foundation, either version 3 of the License, or (at your option) any later
version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY
WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with
this program. If not, see <https://www.gnu.org/licenses/>.
'''

import openpyxl, os, sys

# Variables.
XLSX_FILE = 'tls_caves_photos_master_v1.xlsx'

if __name__=='__main__':
    # Usage.
    if len(sys.argv) < 2:
        sys.stderr.write('Usage:\n')
        sys.stderr.write('python {} --dry-run\n'.format(sys.argv[0]))
        sys.stderr.write('python {} --rename\n'.format(sys.argv[0]))
        sys.exit()

    wb = openpyxl.load_workbook(XLSX_FILE)
    ws = wb.active
    
    abs_paths = []
    for i, row in enumerate(ws.iter_rows()):
        # Skip the header row.
        if i == 0:
            continue
        # Assume that column A contains new filenames, column B contains old
        # filenames, and column C contains the path to each file. Append tuples
        # to the abs_path list. Each tuple contains two elements: an absolute
        # path to the old filename, and an absolute path to the new filename.
        abs_paths.append(
            (
                os.path.join(row[2].value, row[1].value),
                os.path.join(row[2].value, row[0].value)
            )
        )

    # Confirm that old filenames and new filenames each occur only once.
    def c(path_list):
        counts = {}
        for p in path_list:
            if not p in counts:
                counts[p] = 0
            counts[p] += 1
        for p, count in counts.items():
            if count > 1:
                sys.stderr.write('{} occurs more than once.\n'.format(p))
                sys.exit()
    c([p[0] for p in abs_paths])
    c([p[1] for p in abs_paths])
    
    # Confirm that all of the old filenames exist on disk and are not symlinks.
    # Please note that this script does not check to see if two hardlinks point to
    # the same file.
    for p in abs_paths:
        if not os.path.isfile(p[0]):
            sys.stderr.write('{} is not a file.\n'.format(p[0]))
            sys.exit()
        if os.path.islink(p[0]):
            sys.stderr.write('{} is a symlink.\n'.format(p[0]))
            sys.exit()
        if os.path.isfile(p[1]) or os.path.isdir(p[1]):
            sys.stderr.write('{} already exists on disk.\n'.format(p[1]))
            sys.exit()

    if sys.argv[1] == '--dry-run':
        for p in abs_paths:
            sys.stdout.write('rename "{}" to "{}"\n'.format(p[0], p[1]))
    elif sys.argv[1] == '--rename':
        for p in abs_paths:
            os.rename(p[0], p[1])
