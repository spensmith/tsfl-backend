import time

import numpy as np
import pandas as pd

from autograder import TESTING


def potential_sleep(sleep_seconds):
    time.sleep(0 if TESTING is True else sleep_seconds * 0.5)


def empty_string_to_null(input_object):
    if pd.isna(input_object):
        return np.nan
    elif str(input_object).lower() in ('', 'nan', 'nat', 'none'):
        return np.nan
    elif isinstance(input_object, str) and any([input_object.isspace(), not input_object]):
        return np.nan
    elif input_object is None:
        return np.nan
    return input_object


def format_excel_worksheet(worksheet, dataframe):
    for i, col in enumerate(list(dataframe)):
        iterate_length = dataframe[col].astype(str).str.len().max()
        header_length = len(col)
        max_size = max(iterate_length, header_length) + 1
        worksheet.set_column(i, i, max_size)


def conditional_format(worksheet, workbook, column_format_range, winning_number_of_games):
    if workbook:
        colors_dictionary = {
            '0': {
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006'
            },
            winning_number_of_games: {
                'bg_color': '#C6EFCE',
                'font_color': '#006100'
            }
        }
        for if_equals, format_dictionary in colors_dictionary.items():
            excel_format = workbook.add_format(format_dictionary)
            worksheet.conditional_format(column_format_range, {
                'type': 'cell',
                'criteria': '=',
                'value': if_equals,
                'format': excel_format
            })


def remove_inbetween_quotations(name):
    try:
        index_for_first_quotation = name.find('"')
        index_for_second_quotation = name.find('"', index_for_first_quotation + 1)
        return name[:index_for_first_quotation] + name[index_for_second_quotation + 1:]
    except Exception:
        return name


def remove_inbetween_open_and_close_paren(name):
    try:
        index_for_open_paren = name.find('(')
        index_for_close_paren = name.find(')', index_for_open_paren + 1)
        return name[:index_for_open_paren] + name[index_for_close_paren + 1:]
    except Exception:
        return name


def remove_and_following(name, and_phrase):
    try:
        index_for_and = name.find(and_phrase)
        index_for_following = name.find(' ', index_for_and + 1)
        return name[:index_for_and] + name[index_for_following + 1:]
    except Exception:
        return name


def quotation_cleaner(name):
    while '"' in name:
        name = remove_inbetween_quotations(name)
    return name


def paren_cleaner(name):
    while '(' in name and ')' in name:
        name = remove_inbetween_open_and_close_paren(name)
    return name


def and_cleaner(name):
    for and_phrase in (' and ', ' & '):
        while and_phrase in name:
            name = remove_and_following(name, and_phrase=and_phrase)
    return name


def get_first_and_last_with_chars(name, first_name_stub_size, last_name_stub_size, use_first_letter_of_third_word):
    name = str(name).strip()
    for cleaner in (quotation_cleaner, paren_cleaner, and_cleaner):
        name = cleaner(name)

    formatted_name = ''
    name_split = list(filter(None, [word.strip() for word in name.split(' ')]))
    for i, word in enumerate(name_split):
        formatted_name += ' ' if 0 < i < len(name_split) else ''
        if i == 0:
            formatted_name += word[:first_name_stub_size]
        if i == 1:
            formatted_name += word[:last_name_stub_size]
        elif i == 2 and use_first_letter_of_third_word is True:
            formatted_name += word[0]
    return formatted_name.strip()


def get_letter_from_column(dataframe, week_string):
    for i, col in enumerate(list(dataframe)):
        if col == week_string:
            return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[i]


def get_filename_and_sheetname(label):
    if label.endswith('.xlsx'):
        return label, label.split('xlsx')[0]
    else:
        return label + '.xlsx', label


def export_excel(dataframe, label):
    filename, sheetname = get_filename_and_sheetname(label)
    print(f'Now exporting: {filename}')

    with pd.ExcelWriter(filename) as report_writer:
        dataframe.to_excel(report_writer, sheet_name=sheetname, index=False)
        format_excel_worksheet(report_writer.sheets[sheetname], dataframe)


def get_name_iterator():
    for use_first_letter_of_third_word in [True, False]:
        for first_name_stub_size in [4, 3]:
            for last_name_stub_size in [4, 3]:
                yield first_name_stub_size, last_name_stub_size, use_first_letter_of_third_word


def get_current_column_name(week_number, column_names):
    for column_name in column_names:
        if f'{week_number:02}' in column_name and 'week' in column_name:
            return column_name
    return f'Week {week_number:02}'
