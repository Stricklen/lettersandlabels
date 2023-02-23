import tkinter as tk

colour_names_link  = 'https://www.tcl.tk/man/tcl8.5/TkCmd/colors.html'

def update_colour(colour, widget):
    widget['bg'] = colour

root = tk.Tk()

frame = tk.Frame(root)
frame.pack()

entry = tk.Entry(frame)
entry.pack()
text_box = tk.Text(root)
text_box.pack()
update = tk.Button(frame, text='update', command=lambda: update_colour(entry.get(), text_box))
update.pack()

root.mainloop()

