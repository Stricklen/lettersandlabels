import tkinter as tk
from tkinter import ttk, messagebox
import main
import asyncio

class App(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

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
        self.create_widgets()

    def create_widgets(self):
        options_frame = ttk.LabelFrame(self.container, text='Options')
        options_frame.grid(row=0, column=0)

        country_label = ttk.Label(options_frame, text='Country')
        country_label.grid(row=0,column=0)
        self.country = tk.StringVar()
        country_menu = ttk.OptionMenu(options_frame,
                                      self.country,
                                      self.countries[0],
                                      *self.countries)
        country_menu.grid(row=0,column=1)

        company_label = ttk.Label(options_frame, text='Company')
        company_label.grid(row=1, column=0)
        self.company = tk.StringVar()
        company_menu = ttk.OptionMenu(options_frame,
                                      self.company,
                                      self.companies[0],
                                      *self.companies)
        company_menu.grid(row=1, column=1)

        letters_container = tk.LabelFrame(self.container,
                                          text='Letters (seperate names with ,)',
                                          width=30)
        letters_container.grid(row=1,column=0)
        self.name_entry = ttk.Entry(letters_container, width=50)
        self.name_entry.pack()
        submit_btn = ttk.Button(letters_container, text='Print',command=lambda:asyncio.run(self.print_letters()))
        submit_btn.pack()

    async def print_letters(self):
        names_list = self.name_entry.get().split(',')
        await main.print_letters(names_list, str(self.company.get()).lower(), str(self.country.get()).lower())
        messagebox.showinfo(title='SUCCESS!', message='Successfully printing letters')


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    app.pack(side='top', fill='both', expand=True)
    root.mainloop()
