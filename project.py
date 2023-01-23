#  pip install openpyxl
#  pip install xlrd
#  pip install pillow
#  pip install matplotlib
#  pip install ttkbootstrap

from tkinter import ttk
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import ttkbootstrap as ttkb
import os
import openpyxl
import xlrd
from openpyxl import Workbook
import pathlib


background = "#06283D"
framebg = "#EDEDED"
framefg = "#06283D"
root = tk.Tk()
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)

style = ttkb.Style('superhero')
style.configure("TButton", font=("Helvetica", 12, "bold"))
root.title('Our Registration App')
root.iconbitmap('orange.ico')
# root.geometry("1250x700+110+30")
# root.resizable(False,False)
# root.config(bg=background)
file = pathlib.Path('lab.xlsx')

if file.exists():
    pass
else:
    file = Workbook()
    sheet = file.active
    sheet['A1'] = "ID"
    sheet['B1'] = "Name"
    sheet['C1'] = "Family"
    sheet['D1'] = "Postalcode"
    sheet['E1'] = "work experience"
    sheet['F1'] = "salary"

    file.save('lab.xlsx')

input_frame = ttkb.Labelframe(text="input")
input_frame.grid(row=0, column=0)

name_label = ttkb.Label(
    input_frame, text="Name:", bootstyle="success", font=("Helvetica", 12, "bold"))
name_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
name_entry = ttkb.Entry(input_frame, width=25)
name_entry.grid(row=0, column=1, padx=10, pady=10)

family_label = ttkb.Label(
    input_frame, text="Family:", bootstyle="success", font=("Helvetica", 12, "bold"))
family_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
family_entry = ttkb.Entry(input_frame, width=25)
family_entry.grid(row=1, column=1, padx=10, pady=10)

Postalcode_label = ttkb.Label(
    input_frame, text="Postalcode:", bootstyle="success", font=("Helvetica", 12, "bold"))
Postalcode_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
Postalcode_entry = ttkb.Entry(input_frame, width=25)
Postalcode_entry.grid(row=2, column=1, padx=10, pady=10)

work_experience_label = ttkb.Label(
    input_frame, text="work experience:", bootstyle="success", font=("Helvetica", 12, "bold"))
work_experience_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
work_experience_entry = ttkb.Entry(input_frame, width=25)
work_experience_entry.grid(row=3, column=1, padx=10, pady=10)

salary_label = ttkb.Label(
    input_frame, text="salary:", bootstyle="success", font=("Helvetica", 12, "bold"))
salary_label.grid(row=4, column=0, padx=10, pady=10, sticky="w")
salary_entry = ttkb.Entry(input_frame, width=25)
salary_entry.grid(row=4, column=1, padx=10, pady=10)


# buttons frame

button_frame = ttkb.Frame(root)
button_frame.grid(row=1, column=0, pady=10)

save_button = ttkb.Button(button_frame, text="save",
                          bootstyle="success outline")
save_button.grid(row=0, column=0, sticky="ew")

root.mainloop()
