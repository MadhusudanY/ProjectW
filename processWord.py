import docx2txt
import re
from pathlib import Path
from openpyxl import Workbook


# Function to get folder containing position descriptions documents from the console input
def get_path():
    s = input("Please enter folder path : ")
    c = input("Please enter customer name : ").upper()
    return s.replace('\\', '/'), c


def massage_word_data(client_name, str_content):
    if client_name == "FEATHER RIVER COMMUNITY COLLEGE" or client_name == "FRC":
        str_content = re.sub(r'Page [0-9]+ of [0-9]+', '', str_content)
        # print(str_content)
        str_content = str_content.replace("FEATHER RIVER", "").replace("COMMUNITY COLLEGE DISTRICT", "") \
            .replace("570 Golden Eagle Ave., Quincy CA 95971", "").replace("(530) 283-0202, ext. 257", "").replace(
            "www.frc.edu", "").strip()
        return str_content
    elif client_name == "COCCD":
        str_content = re.sub(r'Page [0-9]+ of [0-9]+', '', str_content)
        # print(str_content)
        str_content = str_content.replace("FEATHER RIVER", "").replace("COMMUNITY COLLEGE DISTRICT", "") \
            .replace("570 Golden Eagle Ave., Quincy CA 95971", "").replace("(530) 283-0202, ext. 257", "").replace(
            "www.frc.edu", "").strip()
        return str_content


def get_headers(client_name):
    with open('Headers.txt', 'r') as f:
        for line in f.readlines():
            if line.startswith(client_name):
                headers_list = line.split('|')
                headers_dict = {}
                for x in headers_list[1:]:
                    if x.count('-') == 0:
                        headers_dict[x] = ""
                    else:
                        x = x.split("-")[0]
                        headers_dict[x] = ""
                f.close()
    return headers_dict


def get_chug_it(client_name):
    chug_it = 0
    with open('Headers.txt', 'r') as f:
        for line in f.readlines():
            if line.startswith(client_name):
                chug_it = int(line.split('|')[-1])
                f.close()
                break
    return chug_it * -1


def get_col_data(client, headers, data):
    some_list = list(headers.keys())
    some_list.remove('FILENAME')
    marker1, marker2 = "", ""
    for i in range(len(some_list) - 1):
        marker1 = some_list[i]
        marker2 = some_list[i + 1]
        col_data = ""
        regex_pattern = re.compile(f'^{marker1}([\S\s]*)^{marker2}', re.MULTILINE)
        if regex_pattern.search(data) is not None:
            col_data = regex_pattern.search(data).group().replace(marker1, "").replace(marker2, "").strip()
        elif regex_pattern.search(data) is None:
            if data.lower().count(marker1.lower()) > 0 and data.lower().count(marker2.lower()) > 0:
                start_index = data.lower().index(marker1.lower()) + len(marker1)
                end_index = data.lower().index(marker2.lower())
                col_data = data[start_index:end_index].strip()
        headers[marker1] = col_data
    if data.lower().count(marker2.lower()) > 0:
        start_index = data.lower().index(marker2.lower()) + len(marker2)
        chug_it = get_chug_it(client)
        if chug_it < 0:
            col_data = data[start_index:chug_it].strip()
        else:
            col_data = data[start_index:].strip()
        headers[marker2] = col_data
    return headers


def modify_operation(value, modifiers, dict_headers):
    for item in modifiers:
        if item.startswith('replaceWithEmpty'):
            a, b = item.split(' ')
            value = value.replace(b, '')
        if item.startswith('sliceLastIndexOf'):
            a, b, c = item.split(' ')
            c = int(c)
            if c == 0:
                c = -1
            value = value[len(value) - value[::-1].index(b):]
        if item.startswith('replaceAsContains'):
            a, b, c = item.split(' ')
            if value.lower().__contains__(b.lower()):
                value = b
            elif value.lower().__contains__(c.lower()):
                value = c
        if item.startswith('replaceFromKey'):
            key = item[item.index('(')+1:item.index(')')]
            value = dict_headers.get(key)
            a, b = item.replace(f'replaceFromKey({key}) ', '').split('-')
            if a == 'AlphaNumericOneChar':
                value = re.findall("[a-zA-Z0-9]", value)[int(b)]
    return value.strip()


def apply_modifiers(client, dict_headers):
    modifiers_list = []
    with open('Modifiers.txt', 'r') as f:
        for line in f.readlines():
            if line.startswith(client):
                modifiers_list = line.split('|')[1:]
                f.close()
                break
    for item in modifiers_list:
        col_header = item.split('#')[0]
        modifiers = item.split('#')[1:]
        dict_headers[col_header] = modify_operation(dict_headers.get(col_header), modifiers, dict_headers)
    return dict_headers


def execute(tb=None):
    pd_dir, client = get_path()
    directory = Path(pd_dir).glob("*.docx")
    wb = Workbook()
    xl_file = f"{pd_dir}/PositionDescriptions.xlsx"
    ws = wb.create_sheet(index=0, title="PositionDescriptions")
    row = 2
    try:
        for my_file in directory:
            word_data = docx2txt.process(my_file)
            # This is for debugging Purpose
            print(word_data)
            # exit(0)
            dict_headers = get_headers(client)
            dict_headers['FILENAME'] = my_file.name
            # Last line Item on Header.txt is the string value that needs to be excluded. I called it chug_it value.
            # Hence pop it from dict_headers
            dict_headers.popitem()
            dict_headers = get_col_data(client, dict_headers, word_data)
            # This is for debugging Purpose
            print(dict_headers)
            dict_headers = apply_modifiers(client, dict_headers)
            try:
                col = 1
                for k in dict_headers.keys():
                    ws.cell(1, col).value = k.replace(":", "")
                    ws.cell(row, col).value = dict_headers.get(k)
                    col += 1
                row += 1
            except IndexError:
                print("Something went wrong with " + my_file.name)
                continue
        print("Execution Completed")
    except PermissionError as p:
        print("Permission denied. Please see the docx file(s) are closed.")
        raise Exception(print("Something went wrong :( "))
    finally:
        wb.save(xl_file)
