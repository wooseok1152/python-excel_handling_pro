import tkinter as tk
from tkinter import filedialog
from PIL import ImageTk, Image
import pandas as pd

fileDirectory = ""

def address():

    global fileDirectory1
    fileDirectory1 = filedialog.askopenfilename(initialdir="C:/", title="엑셀파일 선택", filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

def process_excel():

    try :

        sheetName = entry1.get()
        attributeName = entry2.get()
        targetString = entry3.get()
        targetString = set(targetString.replace(" ", "").upper())
        newString = entry4.get()
        df = pd.read_excel(fileDirectory1, sheet_name = sheetName)
        list = df[attributeName]
        for i in range(len(list)):

            list[i] = list[i].replace(" ", "").upper()
            set_element = set(list[i])
            if targetString == set_element :

                list[i] = newString

        fileDirectory2 = filedialog.askdirectory();
        fileDirectory2 = fileDirectory2 + "/new.xlsx"
        df.to_excel(fileDirectory2, index = False)
        result_label.config(text="Complete!", anchor="center")

    except Exception :

        result_label.config(text="Fail! Check the textbox!", anchor="center")

window = tk.Tk()
window.title("For Lovelyn")
window.geometry("700x600+%d+150" % (660))
window.resizable(False, False)

leftFrame = tk.Frame(window)
leftFrame.pack(side="left", anchor="n", pady=5)

frame1 = tk.Frame(leftFrame, width=50)
frame1.pack(anchor="w")

file_button = tk.Button(frame1, text="파일 선택", width=27, relief = "solid", bd = 2, command=address)
file_button.pack(anchor="center")

frame2 = tk.Frame(leftFrame)
frame2.pack(anchor="w")

sheet_label = tk.Label(frame2, text="시트명 : ")
sheet_label.grid(row=0, column=0)

entry1 = tk.Entry(frame2)
entry1.grid(row=0, column=1)

frame3 = tk.Frame(leftFrame)
frame3.pack(anchor="w")

attribute_label = tk.Label(frame3, text="필드명 : ")
attribute_label.grid(row=0, column=0)

entry2 = tk.Entry(frame3)
entry2.grid(row=0, column=1)

frame4 = tk.Frame(leftFrame)
frame4.pack(anchor="w", pady = 10)

target_string_label = tk.Label(frame4, text="타켓 문자열 : ", width = 27, anchor = "center")
target_string_label.pack()

entry3 = tk.Entry(frame4, width = 27)
entry3.pack(pady = 2)

changed_string_label = tk.Label(frame4, text="변형 문자열 : ", width = 27, anchor = "center")
changed_string_label.pack()

entry4 = tk.Entry(frame4, width = 27)
entry4.pack(pady = 2)

frame5 = tk.Frame(leftFrame)
frame5.pack(anchor="w")

file_button = tk.Button(frame5, text="클릭!", width=27, height = 5, relief = "solid", bd = 2, command=process_excel)
file_button.pack(anchor="center")

result_label = tk.Label(leftFrame, width=27)
result_label.pack(anchor="w")

photo = ImageTk.PhotoImage(file = ".img/unnamed2.jpg")
rightLabel = tk.Label(window, relief="solid", width=600, height=600, image = photo)
rightLabel.pack(side="left", anchor="n", padx=5, pady=5)

window.mainloop()

