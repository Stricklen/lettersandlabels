from docxtpl import DocxTemplate
from win32com import client
import pathlib
import asyncio
import os
import pythoncom

cur_path = pathlib.Path(__file__).parent.resolve()

word = client.Dispatch("Word.Application")

boxes = 6
rows = 10

letter_queue = []
letter_queue_active = False

i = 0


def run_queue():
    global letter_queue_active, letter_queue
    if letter_queue_active:
        return
    letter_queue_active = True
    while letter_queue:
        address_layout, company, progress_window, country  = letter_queue[0]
        asyncio.run(print_letters(address_layout, company, progress_window, country))
        letter_queue.pop(0)
    letter_queue_active = False


def add_to_queue(item):
    global letter_queue
    letter_queue.append(item)
    run_queue()


def fill_label_context(dict_in):
    for i in range(1,boxes+1):
        for j in range(1,rows+1):
            key = f'addressx{i}x{j}'
            if key not in dict_in:
                dict_in[key]=''
    return dict_in


async def print_file(file_path):
    pythoncom.CoInitialize()
    word = client.Dispatch("Word.Application")
    word.Documents.Open(file_path)
    word.ActiveDocument.PrintOut()
    word.ActiveDocument.Close()


def get_global_path(local_path):
    global_path = str(cur_path) + local_path
    return global_path.replace('/',r'\^').replace('^','')


async def print_letters(names_list, company, progress_window, country='',):
    # find document for template
    list_len = len(names_list)
    if list_len == 0:
        return False
    increment = 100/list_len
    progress_window.started()

    doc_path = f'./templates/{company}-letter.docx'
    if country:
        doc_path = doc_path.replace('.',f'-{country}.')

    doc = DocxTemplate(doc_path)  # open template

    i = 0
    for name in names_list:
        progress_window.set_status(f'{country.upper()} {company.upper()} - Printing letter for: {name}')
        context = {'name':name}
        doc.render(context)
        file_name = f'{company}-temp-{names_list.index(name)}.docx'
        local_path = f'./temp_files/{file_name}'
        doc.save(local_path)
        abs_path = get_global_path(local_path[1:])
        await print_file(abs_path)
        os.remove(abs_path)
        i+=1
        progress_window.set_progress(increment * i)
    progress_window.destroy()
    return True


def str_to_labeldict(address, box_number):
    split_by_lines = address.split('\n')
    if len(split_by_lines) > boxes:
        raise IndexError
    out_dict = {}
    for line in split_by_lines:
        out_dict[f'addressx{box_number}x{split_by_lines.index(line)+1}'] = line
    return out_dict


def del_temps():
    all_items = os.listdir('./temp_files')
    temps_path = get_global_path('/temp_files/')
    for item in all_items:
        my_path = temps_path + item
        os.remove(my_path)


if __name__ == '__main__':
    print('attempting to run main.py, try running app.py instead.')
