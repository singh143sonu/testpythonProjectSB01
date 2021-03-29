import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

root = tk.Tk()

canvas1 = tk.Canvas(root, width=300, height=300, bg='lightsteelblue2', relief='raised')
canvas1.pack()

label1 = tk.Label(root, text='File Conversion Tool', bg='lightsteelblue2')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)
read_file_excel = 0
read_file_csv = 0
import_file_path = 0
type_of_file_csv = 0
type_of_file_excel = 0

def getFile():

    global read_file_excel
    global read_file_csv
    global import_file_path
    global type_of_file_excel
    global type_of_file_csv

    import_file_path = filedialog.askopenfilename()
    type_of_file_csv = import_file_path[:-4:-1]         # it should start with vsc.
    type_of_file_excel = import_file_path[:-5:-1]       # it should start with xslx.
    if type_of_file_excel == 'xslx':
        read_file_excel = pd.read_excel(import_file_path)
    elif type_of_file_csv == 'vsc':
        read_file_csv = pd.read_csv(import_file_path)
    else:
        MsgBox = tk.messagebox.showinfo('Incorrect file selection', 'Please select a valid file type. Either .csv or .xlsx', icon='warning')

#----------------------------
def convertToFile():
    # if  import_file_path is None:
    #     MsgBox = tk.messagebox.showinfo('No file selected',
    #                                     'No file selected. Please select a valid file of .csv or .xlsx type.',
    #                                     icon='warning')
    if type_of_file_excel == 'xslx':
        export_file_path = filedialog.asksaveasfilename(defaultextension='.csv')
        read_file_excel.to_csv(export_file_path, index=None, header=True)
    elif type_of_file_csv == 'vsc':
        export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        read_file_csv.to_excel(export_file_path, index=None, header=True)

def exitApplication():
    MsgBox = tk.messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application',
                                       icon='warning')
    if MsgBox == 'yes':
        root.destroy()
#-------------------------------
browseButton_Excel = tk.Button(text="     Import Excel or CSV File     ", command=getFile, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 100, window=browseButton_Excel)
#-------------------------------
saveAsButton_CSV = tk.Button(text='Convert File to CSV or Excel', command=convertToFile, bg='green', fg='white',
                             font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=saveAsButton_CSV)
#-------------------------------
exitButton = tk.Button(root, text='       Exit Application     ', command=exitApplication, bg='brown', fg='white',
                       font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 250, window=exitButton)
#-------------------------------

root.mainloop()