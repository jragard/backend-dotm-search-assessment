#!/usr/bin/env python
"""
Given a directory path, this searches all files with extension .dotm in the path,
for text string within the 'word/document.xml' section of a MSWord .dotm file.
"""

import os
import zipfile
import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dir', help='directory path containing dotm files to search. Default is cwd.')
    parser.add_argument('text', help='text to search within each dotm file')
    args = parser.parse_args()

    search_text = args.text
    search_path = args.dir

    if not search_text:
        parser.print_usage()
        exit(1)

    if search_path is None:
        search_path = '.'

    print("Searching directory {} for text '{}' ...").format(search_path, search_text)

    # Get a list of all the files in the search path.
    # This is not a recursive search.
    # Could also use os.walk()
    file_list = os.listdir(search_path)
    DOC_FILENAME = 'word/document.xml'
    match_count = 0
    search_count = 0

    # Iterate over each file in search path
    for file in file_list:

        # Don't care about other file extensions besides dotm
        if not file.endswith('.dotm'):
            continue
        else:
            search_count += 1

        # Construct the full file path
        full_path = os.path.join(search_path, file)

        # Is this dotm file a zip archive format?
        if zipfile.is_zipfile(full_path):

            with zipfile.ZipFile(full_path) as z:
                # Get table of contents
                names = z.namelist()
                # we are interested in the specific doc named 'word/document.xml' 'r'=ReadOnly
                if DOC_FILENAME in names:
                    with z.open(DOC_FILENAME, 'r') as doc:
                        # doc now contains xml.
                        # We could parse all the xml as well ... do we really need to?
                        # Just read the xml contents line by line, looking for our search_text
                        for line in doc:
                            text_location = line.find(search_text)
                            if text_location >= 0:
                                match_count += 1
                                print('Match found in file {}'.format(full_path))
                                print('   ...' + line[text_location-40:text_location+40] + '...')

    print('Total dotm files searched: {}'.format(search_count))
    print('Total dotm files matched: {}'.format(match_count))


if __name__ == '__main__':
    main()




