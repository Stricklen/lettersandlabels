import tkinter as tk
from tkinter import ttk, messagebox, font
import threading
import main
import asyncio
import address_pdf as ad_pdf
import time


queue_active = False
queue_list = []


def run_queue():
    global queue_active, queue_list
    if queue_active:
        return
    queue_active = True
    while queue_list:
        if queue_list[0][0] == 'address':
            address_item = queue_list[0][1]
            address_layout, progress_window = address_item
            ad_pdf.print_address_labels(address_layout, progress_window)
            queue_list.pop(0)
        elif queue_list[0][0] == 'letter':
            letter_item = queue_list[0][1]
            names_list, company, progress_window, country  = letter_item
            asyncio.run(main.print_letters(names_list, company, progress_window, country))
            queue_list.pop(0)
    queue_active = False


def queue_item(item):
    global queue_list
    queue_list.append(item)
    run_queue()


letter_history = []
address_history = [
    [{'company': 'nrs', 'address': '1', 'coord': (0, 0)},
     {'company': 'nrs', 'address': '2', 'coord': (0, 1)},
     {'company': 'nrs', 'address': '3', 'coord': (1, 0)},
     {'company': 'nrs', 'address': '4', 'coord': (1, 1)},
     {'company': 'nrs', 'address': '5', 'coord': (2, 0)},
     {'company': 'nrs', 'address': '6', 'coord': (2, 1)}],
    [{'company': 'nrs', 'address': '7', 'coord': (0, 0)},
     {'company': 'nrs', 'address': '8', 'coord': (0, 1)},
     {'company': 'nrs', 'address': '9', 'coord': (1, 0)},
     {'company': 'nrs', 'address': '10', 'coord': (1, 1)},
     {'company': 'nrs', 'address': '11', 'coord': (2, 0)},
     {'company': 'nrs', 'address': '12', 'coord': (2, 1)}],
    [{'company': 'nrs', 'address': '13', 'coord': (0, 0)},
     {'company': 'nrs', 'address': '14', 'coord': (0, 1)},
     {'company': 'nrs', 'address': '15', 'coord': (1, 0)},
     {'company': 'nrs', 'address': '16', 'coord': (1, 1)},
     {'company': 'nrs', 'address': '17', 'coord': (2, 0)},
     {'company': 'nrs', 'address': '18', 'coord': (2, 1)}],
]


def loadPreviousWindow():
    new_window = tk.Toplevel(root)
    new_window.title('new window')
    print(address_history)
    tk.Label(new_window, text='This is a new window').pack()


class AddressHistoryTable(ttk.Treeview):
    def __init__(self, parent, *args, **kwargs):
        columns = tuple([f'address{n}' for n in range(6)])

        ttk.Treeview.__init__(self, parent, columns=columns, show='headings', *args, **kwargs)


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

        load_previous_btn = ttk.Button(btn_frame, text='Load Previous', command=lambda: loadPreviousWindow())
        load_previous_btn.grid(row=0, column=2)

    def clear_form(self):
        for item in self.boxes.winfo_children():
            item.clear_box()

    def submit_form(self):
        progress_frame = self.master.master.master.nametowidget('.!app.progress_frame')
        progress = ProgressFrame(progress_frame)
        list_out = []
        for item in self.boxes.winfo_children():
            list_out.append(item.get_info())
        progress.pack()
        global address_history
        address_history.append(list_out)
        progress.pack()
        progress_frame.update()
        queue_item(('address', (list_out, progress)))


    def threading_form(self):
        t1 = threading.Thread(target=self.submit_form)
        t1.start()

    def load_previous(self):
        global address_history
        print(address_history)
        messagebox.askquestion(title='Ello', message='Loading previous address')


class AddressBox(tk.Frame):
    def __init__(self, parent, coord, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.coord = coord
        self.variables = ['nrs', 'crs', 'rc', 'nhr']
        self.selection = tk.StringVar()
        self.selection.set('nrs')

        self.address_box = tk.Text(self, height=10, width=40, wrap='none', undo=True, maxundo=10)
        self.address_box['font'] = 'Helvetica 14 bold'

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
                                    command=lambda: self.update_bg())
        crs_check.grid(row=0, column=2)
        rc_check = ttk.Radiobutton(self,
                                   text='RC',
                                   variable=self.selection,
                                   value='rc',
                                   command=lambda: self.update_bg())
        rc_check.grid(row=0, column=3)
        nhr_check = ttk.Radiobutton(self,
                                    text='NHR',
                                    variable=self.selection,
                                    value='nhr',
                                    command=lambda: self.update_bg())
        nhr_check.grid(row=0, column=4)

        self.address_box.grid(row=1, column=0, columnspan=5)

    def enable_entry(self):
        self.address_box.config(state='normal')
    def disable_entry(self):
        self.address_box.config(state='disabled', bg='black', fg='black')

    def update_bg(self):
        selection = self.selection.get()
        if str(selection) == 'nrs':
            self.address_box['bg'] = '#CA3092'
            self.address_box['fg'] = 'white'
            self.enable_entry()
            return
        if str(selection) == 'crs':
            self.address_box['bg'] = '#DDDDDD'
            self.address_box['fg'] = 'black'
            self.enable_entry()
            return
        if str(selection) == 'rc':
            self.address_box['bg'] = 'white'
            self.address_box['fg'] = '#777777'
            self.enable_entry()
            return
        if str(selection) == 'nhr':
            self.address_box['bg'] = 'white'
            self.address_box['fg'] = 'black'
            self.enable_entry()
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
        progress_frame = self.master.master.master.nametowidget('.!app.progress_frame')
        progress = ProgressFrame(progress_frame)
        names_entry = self.name_entry.get('1.0', 'end-1c').split('\n')
        names_list = [name.strip().capitalize() for name in names_entry if name]
        company = str(self.company.get())
        country = str(self.country.get())
        progress.pack()
        progress_frame.update()
        queue_item(('letter',(names_list, company.lower(), progress, country.lower(),)))

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

        self.progress_frame = tk.Frame(self, name='progress_frame')
        self.progress_frame.pack()



if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    app.pack(side='top', fill='both', expand=True)
    root.mainloop()
