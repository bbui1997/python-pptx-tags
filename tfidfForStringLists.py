# Source: http://stevenloria.com/finding-important-words-in-a-document-using-tf-idf/
from __future__ import division, unicode_literals
import math
from textblob import TextBlob as tb
import operator
from time import time
from collections import Counter
# pip install -U textblob
# python -m textblob.download_corpora

# IMPORTANT: In our case, a text is a combination of all the words in a slide
# A list of texts is a list of a combination of all the words in a slide for all the slides in a PowerPoint

# Divide the amount of times a word appears in a giant string of text by
# how many words there are in that string of text
def tf(word, text):
    return float((text.words.count(word))) / float(len(text.words))

# Sums up number of documents that contain that word
def n_containing(word, listOfTexts):
    sum = 0
    for text in listOfTexts:
      if word in text.words:
          sum += 1
    return sum

# log (amount of texts/(1 + how many times word appears in all texts))
def idf(word, listOfTexts):
    return float(math.log(len(listOfTexts))) / float((1 + n_containing(word, listOfTexts)))

# (word count in text / amount of words in text) * log (amount of texts/ 1 + how many times word appears in all texts)
def tfidf(word, text, listOfTexts):
    return float((tf(word, text))) * float(idf(word, listOfTexts))

# sort it by descending order of tf-idf values
# each text has its own top 5, across all slides, count each top 5 and then sort it
# by highest count, return a sorted list of most popular top 5 across all slides
def sortedTfIdfLists(listOfTexts):
    topFive = Counter()
    # for all slides
    for i, text in enumerate(listOfTexts):
        wordCheck = []
        scores = {}
        # figure out tf-idf value for each word in slide
        for word in text.words:
            if(word in wordCheck):
                continue
            scores[word] = tfidf(word, text, listOfTexts)
            wordCheck.append(word)

        # sort it by highest idf and the top 5 to a dict to keep its count
        sorted_words = sorted(scores.items(), key=operator.itemgetter(1), reverse=True)
        for word in sorted_words[:5]:
            item = word[0].encode('ascii', 'ignore')
            topFive[item] += 1

    sortedtopFive = sorted(topFive.items(), key = operator.itemgetter(1), reverse=True)
    return sortedtopFive
