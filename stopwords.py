from stop_words import get_stop_words # pip install stop-words
def filter_stop(words):
    stop_list = get_stop_words('en')
    list = []
    for word in words :
        if word.lower() not in stop_list:
            list.append(word)

    return list
