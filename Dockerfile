FROM python:2.7
ADD pptParser.py /
ADD stopwords.py /
ADD tfidfForStringLists.py /
RUN pip install python-pptx
RUN pip install -U textblob
RUN python -m textblob.download_corpora
CMD [ "python", "./pptParser.py"]