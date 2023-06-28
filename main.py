import tkinter as tk
from tkinter import ttk
from tkinter import font
import openpyxl


def load_data(): # Load data from Excel sheet into treeview widget in GUI using openpyxl module 
    path = "people.xlsx" # Path to Excel sheet 
    workbook = openpyxl.load_workbook(path) # Load workbook that contains Excel sheet 
    sheet = workbook.active # Get active sheet in workbook for reading and writing 

    list_values = list(sheet.values) # Convert sheet values into list of tuples 
    print(list_values)
    for col_name in list_values[0]: # Insert column names into treeview from list of tuples 
        treeview.heading(col_name, text=col_name) # Insert column names into treeview 

    for value_tuple in list_values[1:]: # Insert values into treeview from list of tuples
        treeview.insert('', tk.END, values=value_tuple) # Insert values into treeview 


def insert_row(): # Insert row into Excel sheet and treeview widget in GUI
    name = name_entry.get() # Get name from name entry widget 
    age = int(age_spinbox.get()) # Get age from age spinbox widget
    attending_status = status_combobox.get() # Get attending status from attending status combobox widget
    employment_status = "Employed" if a.get() else "Unemployed"

    print(name, age, attending_status, employment_status)

    # Insert row into Excel sheet
    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, attending_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)
    
    # Clear the values
    name_entry.delete(0, "end") # Clear name entry widget
    name_entry.insert(0, "Name") 
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])

def delete_row():
    selected_item = treeview.selection()[0]
    treeview.delete(selected_item)







def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "Forest-ttk-theme-master/forest-light.tcl")
root.tk.call("source", "Forest-ttk-theme-master/forest-dark.tcl")
root.title("Party RSVP")
root.geometry("800x600")

style.theme_use("forest-dark")

combo_list = ["Attending", "Not Attending", "Other"] # List of options for attending status combobox widget

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Guest List", padding=(20, 10), labelanchor="n")
widgets_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100)
age_spinbox.insert(0, "Age")
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2, column=0, padx=5, pady=5,  sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Attended", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(
    widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("Name", "Age", "Attending", "Present") # Columns are the same as the column names in the Excel sheet and if we want to change the column names in the Excel sheet, we have to change the column names here as well
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("Name", width=100)
treeview.column("Age", width=50)
treeview.column("Attending", width=100)
treeview.column("Present", width=100)
treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()

# Delete button to delete selected row
delete_button = ttk.Button(widgets_frame, text="Delete", command=delete_row)
delete_button.grid(row=7, column=4, padx=5, pady=5, sticky="nsew")



root.mainloop()