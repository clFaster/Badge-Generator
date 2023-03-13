import math
from fpdf import FPDF
import pandas as pd

# CONFIG

# A4 = 297mm x 210mm
a4_width = 297
a4_height = 210

# size in mm
badge_height = 100
badge_width = 75

logo_1_path = "data/l1.png"
logo_2_path = "data/l2.png"
logo_3_path = "data/l3.png"

font_path = "data/Roboto-Regular.ttf"

event_name = "MY SUPER COOL EVENT NAME - 2023"

def main():
    global badge_height, badge_width, event_name

    # badge_height = input_number("Enter badge height in mm (Default 100):", 100)
    # badge_width = input_number("Enter badge width in mm (Default 75):", 75)
    check_valid_size()

    # event_name = input("Enter event name:")

    customers = pd.read_excel(r'data/names.xlsx')

    badges_to_print = len(customers["forename"])

    pages_to_print = math.ceil(badges_to_print / int((a4_width-10) / badge_width))

    print("Pages to print:", pages_to_print)

    badge_number = 0
    pdf_file = FPDF(orientation='L', unit='mm', format='A4')
    pdf_file.add_font('Roboto', '', font_path,
                    uni=True)  # - family, style, fname, uni(Unicode flag)
    pdf_file.set_display_mode(zoom="real")
    pdf_file.set_font("Roboto")

    for page in range(pages_to_print):
        print("Page: ", page + 1)

        pdf_file.add_page()
        draw_cutting_lines(pdf_file)
        boarder_lr = int(287 % badge_width / 2)

        for x in range(5 + boarder_lr, a4_width - boarder_lr - 5, badge_width):
            if badge_number >= badges_to_print:
                break
            pdf_file.rotate(0, a4_width / 2, a4_height / 2)
            draw_badge(pdf_file, x, customers, badge_number)
            pdf_file.rotate(180, a4_width / 2, a4_height / 2)
            draw_badge(pdf_file, a4_width - badge_width - x, customers, badge_number)
            badge_number += 1

    pdf_file = pdf_file.output(name="generated_badge.pdf")


def input_number(message, default_value):
    while True:
        try:
            user_input = int(input(message))
        except ValueError:
            print("Used default.", default_value)
            return default_value
        else:
            return user_input


def check_valid_size():
    if a4_height - 10 < badge_height * 2:
        raise ValueError
    if (a4_width - 10) < badge_width:
        raise ValueError


def draw_cutting_lines(pdf_file, draw_full=False):
    boarder_lr = int(287 % badge_width / 2)
    if draw_full:
        # draw horizontal lines
        for y in range(5, 210, badge_height):
            pdf_file.line(5, y, a4_width-5, y)
        # draw vertical lines
        for x in range(5 + boarder_lr, a4_width, badge_width):
            pdf_file.line(x, 5, x, a4_height-5)
    else:
        for y in range(5, 210, badge_height):
            pdf_file.line(0, y, boarder_lr, y)
            pdf_file.line(297 - boarder_lr - 5, y, 297, y)
        for x in range(5 + boarder_lr, 298 - boarder_lr, badge_width):
            pdf_file.line(x, 0, x, 5)
            pdf_file.line(x, 210 - 5, x, 210)


def draw_badge(pdf_file, x, customers, badge_number):
    # Read Data
    first_name = customers["forename"][badge_number]
    last_name = customers["surname"][badge_number]
    profession = str(customers["profession"][badge_number])
    company = str(customers["company"][badge_number])

    print("Badge: ", badge_number + 1, " - ", first_name, last_name, " - ", profession, " - ", company)

    # Logo 1
    if logo_1_path is not None:
        pdf_file.image(logo_1_path, x=x + 20, y=5 + 2, w=37)

    # Event Name
    pdf_file.set_font_size(size=10)
    pdf_file.set_xy(x=x + 4, y=5 + 25)
    pdf_file.multi_cell(w=badge_width - 6, h=5, txt=event_name, align="L")

    # First Name + Last Name
    pdf_file.set_font_size(size=22)
    pdf_file.text(x=x + 5, y=5 + 47, txt=first_name)
    pdf_file.text(x=x + 5, y=5 + 55, txt=last_name)

    # Profession + Company
    pdf_file.set_font_size(size=14)
    pdf_file.set_xy(x=x + 4, y=5 + 57)
    pdf_file.multi_cell(w=badge_width - 8, h=5, txt=company, align="L")

    # Logo 2
    if logo_2_path is not None:
        pdf_file.set_font_size(size=10)
        pdf_file.text(x=x + 5, y=5 + 72, txt="Hosted by")
        pdf_file.image(logo_2_path, x=x + 3, y=5 + 78, w=33)

    # Logo 3
    if logo_3_path is not None:
        pdf_file.image(logo_3_path, x=x + 42, y=5 + 75, w=27)

if __name__ == '__main__':
    main()
