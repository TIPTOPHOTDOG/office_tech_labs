from docx2pdf import convert
import fitz


def convert_to_pdf(input_file):
    convert(input_file)


def add_signature(input_file, output_file, sign_file):
    sign_rect = fitz.Rect(390, 625, 500, 700)

    file = fitz.open(input_file)
    page = file[0]

    page.insert_image(sign_rect, filename=sign_file)
    file.save(output_file)


convert_to_pdf('../../Downloads/Telegram Desktop/ready.docx')
add_signature('../../Downloads/Telegram Desktop/ready.pdf', 'signed_pdf.pdf', 'sign.png')
