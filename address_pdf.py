import threading
import time
import datetime
import pythoncom
from docxtpl import DocxTemplate
import pathlib
import asyncio
import os
import tkinter as tk
from tkinter import ttk
from docx2pdf import convert
import fitz
from math import ceil


cur_path = pathlib.Path(__file__).parent.resolve()

address_queue = []
address_queue_active = False

i = 0


def run_queue():
    global address_queue_active, address_queue
    if address_queue_active:
        return
    address_queue_active = True
    while address_queue:
        address_layout, progress_window = address_queue[0]
        print_address_labels(address_layout, progress_window)
        address_queue.pop(0)
    address_queue_active = False


def add_to_queue(item):
    global address_queue
    address_queue.append(item)
    run_queue()


def get_global_path(local_path):
    global_path = str(cur_path) + local_path
    return global_path.replace('/',r'\^').replace('^','')


def merge_files(path1, path2):
    doc1 = fitz.open(path1)
    doc2 = fitz.open(path2)

    for i in range(doc1.page_count):
        page = doc1.load_page(i)
        page_front = fitz.open()
        page_front.insert_pdf(doc2, from_page=i, to_page=i)
        page.show_pdf_page(page.rect, page_front, pno=0, keep_proportion=True, overlay=True, oc=0, rotate=0, clip=None)

    doc1.save('result.pdf', encryption=fitz.PDF_ENCRYPT_KEEP)


def merge_pdfs(paths_list, progress_window):
    doc1 = fitz.open(paths_list[0])
    num_paths = len(paths_list)
    if num_paths <1:
        return
    if num_paths <1:
        return paths_list[0]
    increment = 100/num_paths
    progress_window.set_progress(increment)
    for path in paths_list[1:]:
        doc2 = fitz.open(path)
        page = doc1.load_page(0)
        page_front = fitz.open()
        page_front.insert_pdf(doc2)
        page.show_pdf_page(page.rect,
                            page_front,
                            pno=0,
                            keep_proportion=True,
                            overlay=True,
                            oc=0,
                            rotate=0,
                            clip=None)
        progress_window.bump(increment)

    result_path = './temp_files/result.pdf'

    doc1.save(result_path)
    return result_path


def str_to_labeldict(address):
    split_by_lines = [i.strip() for i in address.split('\n')]
    out_dict = {}
    for line in split_by_lines:
        out_dict[f'address{split_by_lines.index(line)+1}'] = line
    return out_dict


def og_convert(path_in, result_path):
    pythoncom.CoInitialize()
    convert(path_in, result_path)


def create_part_pdf(company, coord, address):
    folder = f'./templates/{company}-labels'
    file_path = f'/{coord[0]}-{coord[1]}.docx'
    label_path = folder + file_path
    doc = DocxTemplate(label_path)
    context = str_to_labeldict(address['address'])
    doc.render(context)
    out_path = './temp_files' + file_path
    doc.save(out_path)
    abs_path = get_global_path(out_path)
    og_convert(abs_path, abs_path.replace('.docx', '.pdf'))
    os.remove(abs_path)


def create_pdfs(address_layout, progress_window):
    pdf_paths = []
    increment = 100/len(address_layout)
    for address in address_layout:
        company = address['company']
        coord = address['coord']
        if not company:
            progress_window.bump(increment)
            continue
        address_num = address_layout.index(address)
        progress_window.set_status('Creating address pdf {}'.format(address_num+1))
        create_part_pdf(company, coord, address)
        pdf_path = f'./temp_files/{coord[0]}-{coord[1]}.pdf'
        pdf_paths.append(pdf_path)
        progress_window.set_status('')
        progress_window.bump(increment)
    return pdf_paths


def print_address_labels(address_layout, progress_window):
    address_count = 0
    blanks = []
    for item in address_layout:
        if item['address']:
            address_count += 1
        else:
            blanks.append(item)
    if blanks:
        for blank in blanks:
            address_layout.remove(blank)
    if address_count == 0:
        return False
    progress_window.started()
    progress_window.set_status('Creating address pdfs...')
    pdf_paths = create_pdfs(address_layout, progress_window)

    progress_window.set_progress(0)
    progress_window.set_status('Merging addresses...')
    result_path = merge_pdfs(pdf_paths, progress_window)
    for path in pdf_paths:
        os.remove(get_global_path(path))

    progress_window.set_status('Printing...')
    progress_window.progress.pack_forget()
    os.startfile(get_global_path(result_path), 'Print')
    time.sleep(2)
    progress_window.destroy()


if __name__ == '__main__':
    print('attempting to run address_pdf.py, try running app.py instead.')
