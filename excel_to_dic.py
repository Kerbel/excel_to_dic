import argparse
from tqdm import tqdm
import openpyxl
from collections import defaultdict

NEW_LINE_STRING = '\n'
FIRST_COLUMN_INDEX = 0


def extract_header_data(raw_column):
    header = raw_column.value
    header_index = ''.join(c for c in header if c.isdigit())
    header_string = ''.join(c for c in header if c.isalpha())
    return header_index, header_string


def create_headers_dict():
    headers = {}
    for c in excel_file.columns:
        header_index, header_string = extract_header_data(c[FIRST_COLUMN_INDEX])
        headers[header_index] = header_string
    return headers


def load_words_to_headers_dict(excel_file):
    words_to_header = defaultdict(list)
    for c in excel_file.columns:
        header_index, header_string = extract_header_data(c[FIRST_COLUMN_INDEX])
        for r in [r for r in c[1::] if r.value is not None]:
            words_to_header[r.value].append(header_index)
    return words_to_header


def convert_dict_to_dic_format(words_to_header):
    final_rows = []
    for word, header_indexes in words_to_header.items():
        header_list = "	".join(header_indexes)
        final = f"{word}	{header_list}"
        final_rows.append(final)
    return final_rows


def write_dic_head(file_handle, headers):
    file_handle.write('%')
    file_handle.write(NEW_LINE_STRING)

    for index, header in tqdm(headers.items()):
        file_handle.write(f"{index}	{header}")
        file_handle.write(NEW_LINE_STRING)

    file_handle.write('%')
    file_handle.write(NEW_LINE_STRING)


def write_dic_body(file_handle, word_rows):
    for row in tqdm(word_rows):
        file_handle.write(row)
        file_handle.write(NEW_LINE_STRING)


def write_result_to_file(word_rows, exported_file_path):
    with open(exported_file_path, 'a+') as f:
        print("Write dic headers")
        write_dic_head(f, create_headers_dict())
        print("Write dic body")
        write_dic_body(f, word_rows)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert liwc excel file to .dic file format.')
    parser.add_argument(dest='excel_path', type=str)
    parser.add_argument(dest='output_path', type=str)

    args = parser.parse_args()

    excel_file = openpyxl.load_workbook(args.excel_path).active

    rows = convert_dict_to_dic_format(load_words_to_headers_dict(excel_file))
    write_result_to_file(rows, args.output_path)
