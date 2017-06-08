# TODO:
# 1) Finish this program
#   a. Figure out how we want the user to choose the PowerPoint file
#   b. Error handling involved with choosing a file
#   c. finalize as a group what we want to be our metadata
#   d. figure out/edit how the metadata is being inserted to the program
# 2) Set up code delivery in a way where our user only has to do 1 simple installation to use our program

import codecs
import os
import operator
<<<<<<< HEAD
from Tkinter import *
import Tkinter, Tkconstants, tkFileDialog
=======
from stopwords import filter_stop
>>>>>>> a015a3c11e554fd9da7d283b6579ce10c587df66
from pptx import Presentation # pip install python-pptx

# 1) Take input from user for filename/path
# 2) Parse presentation for text
# 3) Take text and categorize into core properties
# 4) Actually insert those core properties to Presentation

def main():
    prsAndFileNameTuple = findFile()
    prs = prsAndFileNameTuple[0]
    filename = prsAndFileNameTuple[1]

    # Get all of the text or sorted list of most popular words, need to decide
    wordList = parseText(prs)

    # Get 15 item tuple from text, it is possible that some items will be blank
    # author, category, comments, content_status, created, identifier, keywords, language, last_modified_by, last_printed, modified, subject, title, version
    # HOWEVER, we may only have to insert the keywords
    metadata = parseMetaData(wordList)

    # Insert metadata (Core Properties) to appropriate location
    populateCoreProperties(prs, metadata, filename)

def findFile():
    # raw_input() returns a String, input() returns a python expression
    # raw_input() in Python 2.7 is the same as Python3's input()
    # we can figure out how we want the user to input a file name later

    pptx_fileName = tkFileDialog.askopenfilename()
    print (pptx_fileName)
    print()
    prs = Presentation (pptx_fileName)
    #pptx_file.close()
    return prs, pptx_fileName

# Take each slide, read everything that contains a text frame (including shapes)
# Insert it into a list

# for each slide, for each paragraph run, find the runs with the biggest font
# insert each word to a word list
# do a word count of those words in the word list
def parseText(presentation):
    wordList = []
    greatestFontDict = {}
    for slide in presentation.slides:
        max = 0
        greatestRun = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:

                for run in paragraph.runs:
                    # read the text in the text box, encode it from unicode to ascii, get rid of extra white space
                    # then finally add all the contents in the split() list to the word list
                    # could have data loss
                    if(run.font.size > max):
                        greatestRun = run.text.encode('ascii', 'ignore').split()
                        max = run.font.size
        # for word in greatestRun:
        #     print word
        wordList.extend(greatestRun)
    return filter_stop(wordList)

# Figure this out, I don't know how we want to sort out the metadata yet
# for now, all I did was find the top 15 common words used
def parseMetaData(wordList):
    # get the word count
    word_count = {}
    for word in wordList:
        if word in word_count:
            word_count[word] += 1
        else:
            word_count[word] = 1
    sorted_words = sorted(word_count, key=word_count.__getitem__, reverse=True)

    # initialize loop variables
    keywords = ""
    i = 0

    # This if statement is here to get a string with no extra white space
    if len(sorted_words) != 0:
        keywords = sorted_words[0]
        i += 1
    # top 15 or when we run out
    while i < 15 and i < len(sorted_words):
        keywords += ", " + sorted_words[i]
        i += 1

    return keywords


# Add all keywords in wordList to tags property and save
# We may need to change this depending on if we want to add any more metadata
# or how we want to insert it
def populateCoreProperties(presentation, metadata, filename):
    string = ""
    for item in metadata:
        string += item
    presentation.core_properties.keywords = string
    presentation.save(filename)

if __name__ == "__main__":
    main()
