import collections.abc
import os

from bardapi import Bard
from pptx import Presentation
from PyPDF2 import PdfReader

from bard_token import TOKEN


def read_files(input_path: str, output_path: str):
    name_list = os.listdir(input_path)
    return name_list


def ppt2text(input_path: str, name: str):
    prs = Presentation(f"{input_path}/{name}")

    prs_text = ""

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                prs_text = prs_text + paragraph.text + "\n"

    return prs_text


def pdf2text(input_path: str, name: str):
    reader = PdfReader(f"{input_path}/{name}")
    pages = reader.pages

    pdf_text = ""

    for page in pages:
        sub = page.extract_text()
        pdf_text = pdf_text + sub + "\n"

    return pdf_text


def text2bard(text: str, answer: str):
    bard = Bard(token=TOKEN)

    answer = text + answer
    respone = bard.get_answer(answer)["content"]

    return respone


def write_text(name: str, write_text: str, output_path: str):
    name = name.split(".")[0]
    with open(f"{output_path}/{name}.txt", "w", encoding="utf-8") as f:
        f.writelines(write_text)


def run():
    input_path = "./input"
    output_path = "./output"
    answer = "이 텍스트를 요약해줘."

    name_list = read_files(input_path, output_path)

    for name in name_list:
        if name.split(".")[-1] == "pdf":
            text = pdf2text(input_path, name)
            respon = text2bard(text, answer)
            write_text(name, respon, output_path)
        elif name.split(".")[-1] == "pptx":
            text = ppt2text(input_path, name)
            respon = text2bard(text, answer)
            write_text(name, respon, output_path)
        else:
            respon = "error"
            write_text(name, respon, output_path)


if __name__ == "__main__":
    run()
