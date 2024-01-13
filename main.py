from pathlib import Path
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
    for file in files:
        text = extract_text(file)
        d = extract_data_from_text(text)
        for k in d:
            print(f"{k}: {d[k]}")


def extract_data_from_text(text):
    data = dict()
    lines = text.strip().split("\n\n")
    text = "\n".join(lines)

    data['q-id'] = lines[0].split()[-1]
    data['domain'] = lines[8]
    data['skill'] = lines[9]
    data['difficulty'] = lines[-1].split()[-1]
    data['question'] = lines[12]
    re_text = text.replace("\n", "###")

    data['choice-a'] = re.findall("A\\. (.*?)###", re_text)[0].strip()
    data['choice-b'] = re.findall("B\\. (.*?)###", re_text)[0].strip()
    data['choice-c'] = re.findall("C\\. (.*?)###", re_text)[0].strip()
    data['choice-d'] = re.findall("D\\. (.*?)###", re_text)[0].strip()
    data['question'] = re.findall(f"{'ID: ' + data['q-id']}(.*?)A\\.", re_text)[0]
    data['prompt'] = "Which " + re.findall("Which (.*?)\\?", data['question'])[0].replace('###', ' ') + "?"
    data['question'] = re.findall("(.*?)Which", data['question'])[0]
    data['question'] = data['question'].replace('###', '\n').strip()
    data['answer'] = re.findall("Correct Answer: (.*?)###", re_text)[0].strip()
    data['rationale'] = re.findall("Rationale(.*?)" + lines[-1], re_text)[0].replace("###", "\n").strip()

    return data

def write_data_to_excel_sheet(data):
    ...


if __name__ == '__main__':
    # path = get_path()
    path = "C:\\Users\\Niranjan\\CodingWorld\\freelance_work\\pdf_parser\\test\\samples"
    extract_data(path)
