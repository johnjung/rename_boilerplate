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
    
    abs_paths_old = []
    abs_paths_new = []
    
    for i, row in enumerate(ws.iter_rows()):
        # Skip the header row.
        if i == 0:
            continue
        # Assume that column A contains new filenames, column B contains old
        # filenames, and column C contains the path to each file.
        abs_paths_old.append(os.path.join(row[2].value, row[1].value))
        abs_paths_new.append(os.path.join(row[2].value, row[0].value))
    
    # Confirm that old filenames and new filenames each occur only once.
    def c(abs_paths):
        counts = {}
        for f in abs_paths:
            if not f in counts:
                counts[f] = 0
            counts[f] += 1
        for f, count in counts.items():
            if count > 1:
                sys.stderr.write('{} occurs more than once.\n'.format(f))
                sys.exit()
    c(abs_paths_new)
    c(abs_paths_old)
    
    # Confirm that all of the old filenames exist on disk and are not symlinks.
    # Please note that this script does not check to see if two hardlinks point to
    # the same file.
    for f in abs_paths_old:
        if not os.path.isfile(f):
            sys.stderr.write('{} is not a file.\n'.format(f))
            sys.exit()
        if os.path.islink(f):
            sys.stderr.write('{} is a symlink.\n'.format(f))
            sys.exit()

    # Confirm that none of the new filenames exist on disk yet.
    for f in abs_paths_new:
        if os.path.isfile(f) or os.path.isdir(f):
            sys.stderr.write('{} already exists on disk.\n'.format(f))
            sys.exit()

    # Confirm that the list of old pathnames has the same number of items as
    # the list of new path names.
    assert len(abs_paths_old) == len(abs_paths_new)

    if sys.argv[1] == '--dry-run':
        for i in range(len(abs_paths_old)):
            sys.stdout.write('rename "{}" to "{}"\n'.format(
                abs_paths_old[i], 
                abs_paths_new[i]
            ))
    elif sys.argv[1] == '--rename':
        for i in range(len(abs_paths_old)):
            os.rename(abs_paths_old[i], abs_paths_new[i])
