#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!

import zipfile
import sys
import os
import glob



def check_text(text, directory):
    print "Searching directory ./dotm_files for text " + text
    os.chdir(directory)
    filenames = glob.glob('*.dotm')
    matches = 0
    
    
    for file in filenames:     
        zf = zipfile.ZipFile(file, 'r')
        data = zf.read('word/document.xml')
        index = data.find(text)

        # empty strings that we will 'append' the characters before/after the search item
        minus40chars = ""
        plus40chars = ""

        #defines 2 lists of indices representing the 40 characters before and after the index of the search item
        plus_range = range(index+1, index+40)
        minus_range = range(index-40, index-1)

        if text in data:
            matches += 1
            print "Match found in file " + directory + "/" + file
            for index, character in enumerate(data):
                if index in plus_range:
                    plus40chars += character
                if index in minus_range:
                    minus40chars += character

            print "..." + minus40chars + text + plus40chars + "..."

    files_string = str(len(filenames))
    print "Total dotm files searched: " + files_string
    matches_string = str(matches)
    print "Total dotm files matched: " + matches_string

            



def main():
    if len(sys.argv) != 3:
        sys.exit(1)

    return check_text(sys.argv[1], sys.argv[2])
    
    

if __name__ == '__main__':
    main()




