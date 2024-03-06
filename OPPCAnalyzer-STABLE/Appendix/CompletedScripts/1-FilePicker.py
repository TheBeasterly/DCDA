import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

excel_file_path = filedialog.askopenfilename()
