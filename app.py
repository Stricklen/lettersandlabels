import tkinter as tk
from tkinter import ttk, messagebox, font
import threading
import main
import asyncio
import address_pdf as ad_pdf
import time



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

        btn_frame = tk.Frame(self)
        btn_frame.pack()

        submit_btn = ttk.Button(btn_frame, text='Submit', command=lambda: self.threading_form())
        submit_btn.grid(row=0, column=0)

        clear_btn = ttk.Button(btn_frame, text='Clear', command=lambda: self.clear_form())
        clear_btn.grid(row=0, column=1)

    def clear_form(self):
        for item in self.boxes.winfo_children():
            item.clear_box()
    def submit_form(self):
        progress = ProgressFrame(self)
        list_out = []
        for item in self.boxes.winfo_children():
            list_out.append(item.get_info())
        progress.pack()
        ad_pdf.add_to_queue((list_out, progress))

    def threading_form(self):
        t1 = threading.Thread(target=self.submit_form)
        t1.start()


class AddressBox(tk.Frame):
    def __init__(self, parent, coord, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.coord = coord
        self.variables = ['nrs', 'crs', 'rc', 'nhr']
        self.selection = tk.StringVar()
        self.selection.set('nrs')

        self.address_box = tk.Text(self, height=10, width=40, wrap='none', undo=True, maxundo=10)

        self.create_env()

        self.update_bg()


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
                                    command=lambda: self.update_bg())
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

    def update_bg(self):
        selection = self.selection.get()
        if str(selection) == 'nrs':
            self.address_box['bg'] = 'plum1'
            return
        if str(selection):
            self.enable_entry()
            return
        self.disable_entry()


    def get_info(self):
        address = self.address_box.get('1.0', 'end-1c')
        return {'company':self.selection.get(), 'address':address, 'coord':self.coord}

    def clear_box(self):
        self.address_box.delete('1.0', 'end-1c')


class LettersWindow(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)

        self.companies = [
            'NRS',
            'CRS',
            'NHR',
            'RC',
        ]

        self.countries = [
            '',
            'USA',
            'AUS',
        ]

        self.container = tk.Frame(self)
        self.container.pack()
        self.create_env()

    def create_env(self):
        options_frame = ttk.LabelFrame(self.container, text='Options')
        options_frame.grid(row=0, column=0, sticky='n')

        country_label = ttk.Label(options_frame, text='Country')
        country_label.grid(row=0, column=0, sticky='w')
        self.country = tk.StringVar()
        country_menu = ttk.OptionMenu(options_frame,
                                      self.country,
                                      self.countries[0],
                                      *self.countries)
        country_menu.config(width=10)
        country_menu.grid(row=0, column=1, sticky='w')

        company_label = ttk.Label(options_frame, text='Company')
        company_label.grid(row=1, column=0, sticky='w')
        self.company = tk.StringVar()
        company_menu = ttk.OptionMenu(options_frame,
                                      self.company,
                                      self.companies[0],
                                      *self.companies)
        company_menu.config(width=10)
        company_menu.grid(row=1, column=1, sticky='e')

        letters_container = tk.LabelFrame(self.container,
                                          text='Letters: New line for each name',
                                          width=30)
        letters_container.grid(row=0, column=1)
        self.name_entry = tk.Text(letters_container, height=35, width=30, undo=True, maxundo=10)
        self.name_entry.pack()

        btn_frame = tk.Frame(letters_container)
        btn_frame.pack()

        submit_btn = ttk.Button(btn_frame, text='Print', command=lambda: self.threading_letters())
        submit_btn.grid(row=0, column=0)

        clear_btn = ttk.Button(btn_frame, text='Clear', command=lambda: self.clear())
        clear_btn.grid(row=0, column=1)

    def clear(self):
        self.name_entry.delete('1.0', 'end-1c')

    def print_letters(self):
        progress = ProgressFrame(self)
        self.update()
        names_entry = self.name_entry.get('1.0', 'end-1c').split('\n')
        names_list = [name.strip().capitalize() for name in names_entry if name]
        company = str(self.company.get())
        country = str(self.country.get())
        progress.pack()
        main.add_to_queue(
            (names_list,
            company.lower(),
            progress,
            country.lower(),)
        )

    def threading_letters(self):
        t1 = threading.Thread(target=self.print_letters)
        t1.start()


class ProgressFrame(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.length = kwargs.get('bar_length', 300)
        self.default_state='Starting...'
        self.status_label = ttk.Label(self, text=self.default_state)
        self.status_label.pack()
        self.progress = ttk.Progressbar(self, orient='horizontal', length=self.length, mode= 'determinate')

    def reset(self):
        self.pack_forget()
        self.progress['value'] = 0
        self.update_idletasks()
        self.status_label['text']=self.default_state
        self.progress.pack_forget()

    def started(self):
        self.progress.pack()
        self.update()

    def set_status(self, status):
        self.status_label['text']=status

    def set_progress(self, percentage):
        self.progress['value'] = percentage
        self.update_idletasks()

    def bump(self,value):
        self.progress['value'] += value
        if self.progress['value'] > 100:
            self.progress['value'] = 0
        self.update_idletasks()


class App(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.nb = ttk.Notebook(self)

        self.create_env()

    def create_env(self):
        letter_window = LettersWindow(self.nb)
        address_window = AddressWindow(self.nb)

        letter_window.pack()
        address_window.pack()

        self.nb.add(letter_window, text='Letters')
        self.nb.add(address_window, text='Addresses')

        self.nb.pack()


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    app.pack(side='top', fill='both', expand=True)
    root.mainloop()
