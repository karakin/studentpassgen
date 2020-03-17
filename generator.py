# -*- coding: utf-8 -*-
import xlrd
import xlwt
import sys
import random
import string

symbols = [u'абвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ',
           u'abvgdeejzijklmnoprstufhzcssyyyeuaABVGDEEJZIJKLMNOPRSTUFHZCSSYYYEUA']


def cyr_to_latin(name):
    ret = ''
    for c in name:
        ret += symbols[1][symbols[0].index(c)]
    return ret


def generate_password(letters_length=10, numbers_length=2, symbols_length=2):
    letters = string.ascii_letters
    numbers = string.digits
    password_symbols = u'#$%@&?!'
    string_part = ''.join(random.choice(letters) for i in range(letters_length))
    number_part = ''.join(random.choice(numbers) for i in range(numbers_length))
    symbols_part = ''.join(random.choice(password_symbols) for i in range(symbols_length))

    return number_part + string_part + symbols_part


def main(input_file, sheet_index, output):
    book = xlrd.open_workbook(input_file)

    index = int(sheet_index)
    sheet = book.sheet_by_index(index)

    out_book = xlwt.Workbook()
    out_sheet = out_book.add_sheet(f'{sheet.name}_gen')

    out_sheet.write(0, 0, u'ФИО Студента')
    out_sheet.write(0, 1, u'Логин')
    out_sheet.write(0, 2, u'Пароль')
    out_sheet.write(0, 3, u'E-mail')

    for row in range(1, sheet.nrows):
        name = sheet.cell_value(row, 0)
        s = name.split(' ', 3)
        latin_login = cyr_to_latin(s[0] + s[1]).lower()
        password = generate_password()
        out_sheet.write(row, 0, name)
        out_sheet.write(row, 1, latin_login)
        out_sheet.write(row, 2, password)

    out_book.save(output)


if len(sys.argv) < 3:
    sys.stdout.flush()
    exit(1)

main(sys.argv[1], sys.argv[2], sys.argv[3])
