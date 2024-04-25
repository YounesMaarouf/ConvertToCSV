import os
import platform
import subprocess
import sys
import customtkinter as tk
from tkinter import filedialog
from index import convert_excel_to_csv
# from PIL import Image

# browse the folders and get the desired file
def browseFile():

    entry.delete(0, len(entry.get()))

    filepath = filedialog.askopenfilename(title="open xlsx file",filetypes= (("xlsx files","*.xlsx"),("xls files","*.xls*")))

    entry.insert(0, filepath)

# preview the converted file
def preview(file_path):

    filename = os.path.basename(file_path)

    desired_part = filename.split('.')[0] + ".csv"

    if platform.system() == "Windows": #windows
        os.startfile(desired_part)
    elif platform.system() == "Darwin":  # macOS
        subprocess.Popen(["open", desired_part])
    else:  # Linux
        subprocess.Popen(["xdg-open", desired_part])

# convert excel to csv and logics
def convertToCSV():
     
    filepath = entry.get()
    try : 
        convert_excel_to_csv(filepath)

        message.configure(text="the excel file converted to csv successfully", text_color="green")

    except :
        message.configure(text="(make sure you entered a valid path/file.xlsx)", text_color="red")
        print(f'An error occured') 

    finally : 
        entry.delete(0, len(entry.get()))

    if check_var.get() == "on": 
        preview(filepath)


#==============================
# UI 
#==============================

# root app
app = tk.CTk()
app.title("Excel converter")
app.geometry("1500x1200")
tk.set_appearance_mode("system") 
tk.set_default_color_theme("blue") 

# main frame
frame = tk.CTkFrame(master=app, width=1200, height=1000, fg_color="#ffffff")
frame.pack(expand=True, pady=40, padx=20)

# componentes

# example_img_data = Image.open(resource_path("example.png"))

# example_img = tk.CTkImage(light_image=example_img_data, dark_image=example_img_data, size=(1000, 500))


tk.CTkLabel(master=frame, text="UPLOAD THE FILE", font=("Helvetica", 30)).pack(anchor="nw", pady=40, padx=25)

# tk.CTkLabel(master=frame, text="See example :", font=("Helvetica", 20)).pack(anchor="nw", pady=(20, 0), padx=25)

# tk.CTkLabel(master=frame, text="", image=example_img,corner_radius=8).pack(anchor="nw" , pady=10, padx=25)

tk.CTkLabel(master=frame, text="Excel Location", font=("Helvetica", 20)).pack(anchor="nw", pady=(0, 10), padx=25)

# this message will informe the user wether the converting is successful or not
message = tk.CTkLabel(master=frame, text="", font=("Helvetica", 10), text_color='green')
message.pack(anchor="nw", pady=0, padx=25)

############# search container ############
search_container = tk.CTkFrame(master=frame ,width=1000, height=50, fg_color="#F0F0F0")
search_container.pack(side="left", pady=20, padx=25)

button = tk.CTkButton(master=search_container, text="Browse", 
command=browseFile, width=255, height=50, corner_radius=100, hover_color="green", hover=True)
entry = tk.CTkEntry(master=search_container, placeholder_text="D://", width=700, height=50 ,border_color="#2A8C55", border_width=2)

entry.grid(row=0, column=0,padx=10, pady=15)
button.grid(row=0, column=10,padx=10, pady=15)

check_var = tk.StringVar(value="on")
checkbox = tk.CTkCheckBox(search_container, text="open when it's finished",variable=check_var, onvalue="on", offvalue="off")
checkbox.grid(sticky="ew", row=10, column=0, padx=(10, 0), pady=15)
############# search container ############

submit = tk.CTkButton(master=search_container, text="Submit", 
command=convertToCSV, height=50, corner_radius=100, hover_color="green", hover=True)

submit.grid(sticky="ew", row=20, column=0, padx=20, pady=20,)


# preview_button.pack( pady=40, padx=20)

app.mainloop()

#==============================
# END UI 
#==============================