from docxtpl import DocxTemplate
import pathlib
import asyncio
import os
import tkinter as tk
from tkinter import ttk
from docx2pdf import convert
import fitz

cur_path = pathlib.Path(__file__).parent.resolve()

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


def merge_pdfs(paths_list):
    doc1 = fitz.open(paths_list[0])

    if len(paths_list)>1:
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

    result_path = './temp_files/result.pdf'

    doc1.save(result_path)
    return result_path


def str_to_labeldict(address):
    split_by_lines = [i.strip() for i in address.split('\n')]
    out_dict = {}
    for line in split_by_lines:
        out_dict[f'address{split_by_lines.index(line)+1}'] = line
    return out_dict


def create_address_labels(address_layout):
    pdf_paths = []
    for address in address_layout:
        company = address['company']
        coord = address['coord']
        if not company:
            continue
        folder = f'./templates/{company}-labels'
        file_path = f'/{coord[0]}-{coord[1]}.docx'
        label_path = folder + file_path
        doc = DocxTemplate(label_path)
        context = str_to_labeldict(address['address'])
        doc.render(context)
        out_path = './temp_files' + file_path
        doc.save(out_path)
        abs_path = get_global_path(out_path)
        convert(abs_path, abs_path.replace('.docx', '.pdf'))
        pdf_paths.append(out_path.replace('.docx','.pdf'))
        os.remove(abs_path)

    result_path = merge_pdfs(pdf_paths)
    for path in pdf_paths:
        os.remove(get_global_path(path))

    os.startfile(get_global_path(result_path), 'Print')


class AddressWindow(tk.Frame):
    def __init__(self,parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)

        self.boxes = tk.Frame(self)

        self.coords = [
            (0,0),
            (0,1),
            (1,0),
            (1,1),
            (2,0),
            (2,1)
        ]

        self.create_env()

    def create_env(self):
        for coord in self.coords:
            address_box = AddressBox(self.boxes, coord)
            address_box.grid(row=coord[0], column=coord[1], padx=5, pady=5)

        self.boxes.pack()

        submit_btn = ttk.Button(self, text='Submit', command=lambda: self.submit_form())
        submit_btn.pack()

    def submit_form(self):
        list_out = []
        for item in self.boxes.winfo_children():
            list_out.append(item.get_info())

        create_address_labels(list_out)


class AddressBox(tk.Frame):
    def __init__(self, parent, coord, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.coord = coord
        self.variables = ['nrs', 'crs', 'rc', 'nhr']
        self.selection = tk.StringVar()
        self.selection.set('nrs')

        self.address_box = tk.Text(self, height=10, width=40, wrap='none')

        self.create_env()


    def create_env(self):
        none_check = ttk.Radiobutton(self,
                                     text='Disabled',
                                     variable=self.selection,
                                     value='',
                                     command=lambda: self.disable_entry())
        none_check.grid(row=0, column=0)
        nrs_check = ttk.Radiobutton(self,
                                    text='NRS',
                                    variable=self.selection,
                                    value='nrs',
                                    command=lambda: self.enable_entry())
        nrs_check.grid(row=0, column=1)
        crs_check = ttk.Radiobutton(self,
                                    text='CRS',
                                    variable=self.selection,
                                    value='crs',
                                    command=lambda: self.enable_entry())
        crs_check.grid(row=0, column=2)
        rc_check = ttk.Radiobutton(self,
                                   text='RC',
                                   variable=self.selection,
                                   value='rc',
                                   command=lambda: self.enable_entry())
        rc_check.grid(row=0, column=3)
        nhr_check = ttk.Radiobutton(self,
                                    text='NHR',
                                    variable=self.selection,
                                    value='nhr',
                                    command=lambda: self.enable_entry())
        nhr_check.grid(row=0, column=4)

        self.address_box.grid(row=1, column=0, columnspan=5)

    def enable_entry(self):
        self.address_box.config(state='normal', bg='white')
    def disable_entry(self):
        self.address_box.config(state='disabled', bg='light grey')

    def get_info(self):
        address = self.address_box.get('1.0', 'end-1c')
        return {'company':self.selection.get(), 'address':address, 'coord':self.coord}

address_list = [
    {'company': 'nrs', 'address': 'one', 'coord': (0, 0)},
    {'company': 'nrs', 'address': 'two', 'coord': (0, 1)},
    {'company': 'nrs', 'address': 'three', 'coord': (1, 0)},
    {'company': 'nrs', 'address': 'four', 'coord': (1, 1)},
    {'company': 'nrs', 'address': 'five', 'coord': (2, 0)},
    {'company': 'nrs', 'address': 'six', 'coord': (2, 1)}]


if __name__ == '__main__':
    root = tk.Tk()
    window = AddressWindow(root)
    window.pack(side='top', fill='both', expand=True)
    root.mainloop()
