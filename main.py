import collections
from pathlib import Path

import openpyxl
from pdfminer.high_level import extract_text
import re
from openpyxl import Workbook


def get_path():
    print("Hello there! Help me out with parsing your files.\n"
          "Kindly give me the path address of the folder with all the pdf files.")

    input_path = input("Path: ")
    while not Path(input_path).is_dir():
        print("Path doesn't exist. Try again!")
        input_path = input("Path: ")

    output_path = input("\nWhere do you want to store the output file?\nPath: ")
    while not Path(output_path).is_dir():
        print("Path doesn't exist. Try again!")
        output_path = input("Path: ")

    return input_path, output_path


def extract_data(input_path):
    path = Path(input_path)
    files = sorted(path.glob("*.pdf"))
    wb = openpyxl.load_workbook("test/CB SAT RW Question Bank.xlsx")
    sheet = wb.active

    for file in files:
        text = extract_text(file)
        d = extract_data_from_text(text)
        d.append(file.name)
        sheet.append(d)

    wb.save("test/CB SAT RW Question Bank.xlsx")


def extract_data_from_text(text):
    data = dict()
    text = remove_latin_ligatures(text)
    lines = text.strip().split("\n\n")
    text = "\n".join(lines)

    re_text = text.replace("\n", "###")

    print(text + "%%%%%%%%%%%%%%\n", re_text)

    data['q-id'] = lines[0].split()[-1]
    data['domain'] = lines[8]
    data['skill'] = lines[9]
    data['difficulty'] = lines[-1].split()[-1]

    data['question'] = re.findall(f"{'ID: ' + data['q-id']}(.*?)A\\.", re_text)[0]
    data['prompt'] = "Which " + re.findall("Which (.*?)\\?", data['question'])[0].replace('###', ' ') + "?"
    data['question'] = re.findall("(.*?)Which", data['question'])[0]
    data['question'] = data['question'].replace('###', '\n').strip()
    data['choice-a'] = re.findall("A\\. (.*?)###", re_text)[0].strip()
    data['choice-b'] = re.findall("B\\. (.*?)###", re_text)[0].strip()
    data['choice-c'] = re.findall("C\\. (.*?)###", re_text)[0].strip()
    data['choice-d'] = re.findall("D\\. (.*?)###", re_text)[0].strip()
    data['answer'] = re.findall("Correct Answer: (.*?)###", re_text)[0].strip()
    data['rationale'] = re.findall("Rationale(.*?)" + lines[-1], re_text)[0].replace("###", "\n").strip()

    res = []
    for key in (
            'q-id', 'domain', 'skill', 'difficulty', 'question', 'prompt', 'choice-a', 'choice-b', 'choice-c',
            'choice-d',
            'answer', 'rationale'):
        res.append(data[key])

    return res


def write_data_to_excel_sheet(data):
    ...


def remove_latin_ligatures(text):
    liga = {
        'ﬀ': 'ff', 'ﬁ': 'fi', 'ﬂ': 'fl', 'ﬃ': 'ffi', 'ﬄ': 'ffl', 'ﬅ': 'ft', 'ﬆ': 'st', 'Æ': 'AE', 'æ': 'ae', 'Ĳ': 'IJ',
        'ĳ': 'ij', 'Œ': 'OE', 'œ': 'oe'
    }
    s = ""
    c = collections.Counter(text)
    for ch in text:
        if ord(ch) == 0 or ord(ch) == 160:
            s += ' '
            continue
        s += liga.get(ch, ch)

    t = collections.Counter(s)

    for k in c:
        if not k.isalnum():
            print("(", c[k], f"'{k}-{ord(k)}'", ")", end=' ')
    print("8888888888888")
    for k in t:
        if not k.isalnum():
            print("(", t[k], f"'{k}-{ord(k)}'", ")", end=' ')

    return s


if __name__ == '__main__':
    # path = get_path()
    path = "C:\\Users\\Niranjan\\CodingWorld\\freelance_work\\pdf_parser\\test\\samples"
    extract_data(path)
