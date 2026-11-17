import tkinter as tk
from tkinter import ttk
import openpyxl
import sys
import os



def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)




def load_data():
    path = resource_path("Cannabis.xlsx")
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    age_spinbox.config(state="normal")
    age_spinbox.delete(0,tk.END)
    age_spinbox.insert(0, "Age")
    age_spinbox.config(state="readonly")
    gender_combobox.config(state="normal")
    gender_combobox.delete(0,tk.END)
    gender_combobox.insert(0, "Gender")
    gender_combobox.config(state="readonly")
    race_combobox.config(state="normal")
    race_combobox.delete(0,tk.END)
    race_combobox.insert(0, "Race")
    race_combobox.config(state="readonly")
    income_combobox.config(state="normal")
    income_combobox.delete(0,tk.END)
    income_combobox.insert(0, "Income")
    income_combobox.config(state="readonly")
    area_combobox.config(state="normal")
    area_combobox.delete(0,tk.END)
    area_combobox.insert(0, "Area")
    area_combobox.config(state="readonly")
    result.config(state="normal")
    result.delete(0, tk.END)
    result.insert(0,"Total")
    result.config(state="readonly")
    a.set(-1)
    b.set(-1)
    c.set(-1)


    for item in dataview.get_children():
        dataview.delete(item)

    list_values = list(sheet.values)
    for col_name in list_values[0]:
        dataview.heading(col_name, text= col_name)

    for value_tuple in list_values[1:]:
            dataview.insert('',tk.END, values=value_tuple)





def toggle_mode():
    if color_switch.instate(["selected"]):
        style.theme_use("forest-light")
        color_switch.config(text="Dark mode")
    else:
        style.theme_use("forest-dark")
        color_switch.config(text="Light mode")

def search():
    path = resource_path("Cannabis.xlsx")
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)

    path2 = resource_path("SearchResults.xlsx")
    workbook2 = openpyxl.load_workbook(path2)
    sheet2 = workbook2.active
    list_values2 = list(sheet2.values)




    for cell in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=9, values_only=True):
        age = age_spinbox.get()
        user_status = "Y" if c.get() == 1 else "N"
        grower_status = "Y" if b.get() == 1 else "N"
        intend_status = "Y" if a.get() == 1 else "N"
        gender = gender_combobox.get()
        race = race_combobox.get()
        income = income_combobox.get()
        area = area_combobox.get()
        if c.get() == -1:
            user_status = cell[1]
        if b.get() == -1:
            grower_status = cell[2]
        if a.get() == -1:
            intend_status = cell[3]
        if age == "Age":
            age = int(cell[4])
        if gender == "Gender":
            gender = cell[5]
        if race == "Race":
            race = cell[6]
        if income == "Income":
                income = cell[7]
        if area == "Area":
            area = cell[8]
        if (cell[1] == user_status and cell[2] == grower_status and cell[3] == intend_status
                 and cell[4] == int(age) and cell[5] == gender and cell[6] == race and cell[7] == income and cell[8] ==
                 area):

            row_values =[cell[0],user_status,grower_status,intend_status,age,gender,race,income,area]
            sheet2.append(row_values)
            list_values2 = list(sheet2.values)


            for item in dataview.get_children():
                dataview.delete(item)

            for value_tuple in list_values2[1:]:
                dataview.insert('', tk.END, values=value_tuple)
    total = sheet2.max_row
    total2 = sheet.max_row
    result.config(state="normal")
    result.delete(0,tk.END)
    result.insert(0,f"{total-1}/{total2-1}")
    result.config(state="readonly")






root = tk.Tk()
theme_path1 = resource_path("forest-light.tcl")
theme_path2 = resource_path("forest-dark.tcl")
root.title("EXEL Viewer")
style = ttk.Style(root)
root.tk.call("source", theme_path1)
root.tk.call("source", theme_path2)
style.theme_use("forest-dark")
gender_list= ["M", "F"]
race_list = ["Black", "White", "Other", "Mult", "Native"]
income_list = ["0-20", "20-40", "40-60", "60-80", "80-100", "100<"]
area_list = ["City", "Town", "Rural", "Suburbs"]
c = tk.IntVar(value=-1)
b =  tk.IntVar(value=-1)
a =  tk.IntVar(value=-1)


cols = ("Respond", "User","Grower", "Intend", "Age", "Gender", "Race", "Income", "Area")

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text= "Search Row")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)


age_spinbox = ttk.Spinbox(widgets_frame,from_= 0, to= 100)
age_spinbox.grid(row=0,column=0,padx=5, pady=5,sticky="ew")


gender_combobox = ttk.Combobox(widgets_frame, values=gender_list)
gender_combobox.bind("<FocusIn>", lambda e: gender_combobox.delete('0', 'end'))
gender_combobox.grid(row=1,column=0, padx=5, pady=5, sticky="ew")

race_combobox = ttk.Combobox(widgets_frame, values=race_list)
race_combobox.bind("<FocusIn>", lambda e: race_combobox.delete('0', 'end'))
race_combobox.grid(row=2,column=0, padx=5, pady=5, sticky="ew")

income_combobox = ttk.Combobox(widgets_frame, values=income_list)
income_combobox.bind("<FocusIn>", lambda e: income_combobox.delete('0', 'end'))
income_combobox.grid(row=3,column=0, padx=5, pady=5, sticky="ew")

area_combobox = ttk.Combobox(widgets_frame, values=area_list)
area_combobox.bind("<FocusIn>", lambda e: area_combobox.delete('0', 'end'))
area_combobox.grid(row=4,column=0, padx=5, pady=5, sticky="ew")

cannabis_button = ttk.Checkbutton(widgets_frame, text="Cannabis User", variable=c, onvalue=1, offvalue=0)
cannabis_button.grid(row=5,column=0,padx=5, pady=5, sticky="nsew")

cannabis_grower = ttk.Checkbutton(widgets_frame, text="Grower", variable=b, onvalue=1, offvalue=0)
cannabis_grower.grid(row=6,column=0,padx=5, pady=5, sticky="nsew")

cannabis_intend = ttk.Checkbutton(widgets_frame, text="Intend", variable=a, onvalue=1, offvalue=0)
cannabis_intend.grid(row=7,column=0,padx=5, pady=5, sticky="nsew")

result = ttk.Entry(widgets_frame)
result.grid(row=8, column=0)


search_button = ttk.Button(widgets_frame, text="Search", command= search)
search_button.grid(row=9, column=0,padx=5, pady=5, sticky="nsew")

clear_button = ttk.Button(widgets_frame, text="Clear", command= load_data)
clear_button.grid(row=10, column=0,padx=5, pady=5, sticky="nsew")



separator = ttk.Separator(widgets_frame)
separator.grid(row=11, column=0 , padx=(20,10),pady=10, sticky="ew")

color_switch = ttk.Checkbutton(widgets_frame, text="Light Mode",style = "Switch", command=toggle_mode)
color_switch.grid(row=11 , column= 0 , padx=5, pady=10, sticky="nsew")

DataFrame = ttk.Frame(frame)
DataFrame.grid(row=0,column=1,pady=10)
side_Scroll = ttk.Scrollbar(DataFrame)
side_Scroll.pack(side="left", fill="y")




dataview = ttk.Treeview(DataFrame, show='headings',yscrollcommand=side_Scroll.set,columns=cols, height= 13)
dataview.column("Respond", width= 50)
dataview.column("User", width= 50)
dataview.column("Grower", width= 50)
dataview.column("Intend", width= 50)
dataview.column("Age", width= 50)
dataview.column("Gender", width= 50)
dataview.column("Race", width= 50)
dataview.column("Income", width= 50)
dataview.column("Area", width= 60)
dataview.pack()
side_Scroll.config(command=dataview.yview)
load_data()

root.mainloop()