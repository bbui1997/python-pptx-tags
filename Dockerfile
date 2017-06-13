FROM python:2.7
ADD pptParser.py /
ADD stopwords.py /
RUN pip install python-pptx
# RUN apt-get -y install python-tk
CMD [ "python", "./pptParser.py"]