import codecs
from pptx import Presentation # pip install python-pptx

# raw_input() returns a String, input() returns a python expression
# raw_input() in Python 2.7 is the same as Python3's input()
# we can figure out how we want the user to input a file name later
pptx_filename = raw_input("Enter pptx filename: ")

prs = Presentation(pptx_filename)
extract = codecs.open("text.txt", "w", "utf-8")

# text_runs will be populated with a list of strings,
# one for each text run in presentation
text_runs = []
for slide in prs.slides:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text_runs.append(run.text)
    text_runs.append("\n") # separate each slide with a

for elements in text_runs:
    extract.write(elements +'\n')

extract.close()


if __name__ == "__main__":
    main()
