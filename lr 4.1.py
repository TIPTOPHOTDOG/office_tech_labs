from docx2pdf import convert
import fitz
from typing import Dict
import re


def convert_to_pdf(input_file) -> None:
    convert(input_file)


def add_signature(input_file: str, output_file: str, sign_file: str, sign_rect) -> None:
    file = fitz.open(input_file)
    page = file[0]

    page.insert_image(sign_rect, filename=sign_file)
    file.save(output_file)
    print(sign_rect)


def beee(filename:str) -> Dict:
    with open(filename, 'rb') as file:
        pdf = fitz.open(file)

        # Извлекаем текст из документа
        text = ''
        for page_num in range(len(pdf)):
            page = pdf[page_num]
            text = page.get_text("text")
        match = re.search(r'Подпись:\s*(.)|Подпись автора:\s*(.)', text, re.IGNORECASE)
        stat_x = match.start() / 19.5
        return stat_x


convert_to_pdf('Согласие на указание сведений об авторе.docx')
x = beee('Согласие на указание сведений об авторе.pdf')
convert_to_pdf('Согласие на указание сведений об авторе.docx')
x1 = beee('Согласие на указание сведений об авторе.pdf')
add_signature('Согласие на указание сведений об авторе.docx', 'signed_pdf.pdf', 'sign.png', fitz.Rect(x, x + 400, x + 100, x + 500))
add_signature('Согласие на указание сведений об авторе.pdf', 'signed_1234.pdf', 'sign.png', fitz.Rect(x1, x1 + 400, x1 + 100, x1 + 500))
