from os import startfile
from shutil import copyfile
import tkinter as tk
from tkinter.filedialog import askopenfilenames
from openpyxl import load_workbook
from typing import List

class App(tk.Tk):
    def __init__(self):
        super().__init__() 
        self.title('BEFE: Batch Excel File Editor')
        #File list
        self.frm_list = tk.Frame(self)
        self.lbl = tk.Label(self.frm_list, font=("Courier", 18), text="Excel Files List")
        self.listScroll = tk.Scrollbar(self.frm_list, orient=tk.VERTICAL)
        self.file_list = tk.Listbox(self.frm_list, activestyle = tk.DOTBOX, font=('Coursier',20), relief=tk.RAISED, selectmode = tk.SINGLE, yscrollcommand=self.listScroll.set, width=60)
        self.file_list.bind("<Double-Button-1>", self.open_file)
        self.lbl.pack(side=tk.TOP)
        self.file_list.pack(fill=tk.BOTH, expand=True)
        self.frm_list.pack(fill=tk.X, side=tk.TOP, padx=10)


        #file actions buttons
        self.frm_fileAction = tk.Frame(self)
        self.btn_add = tk.Button(self.frm_fileAction, text="Add file", font=('Coursier',16), command=self.add_file)
        self.btn_remove = tk.Button(self.frm_fileAction, text="Remove file", font=('Coursier',16), command=self.remove_file)
        self.btn_deselect = tk.Button(self.frm_fileAction, text="Back up files", font=('Coursier',16), command=self.back_up)
        self.btn_add.grid(row=0,column=0, sticky="ew", padx=5)
        self.btn_remove.grid(row=0,column=1, sticky="ew", padx=5)
        self.btn_deselect.grid(row=0,column=2, sticky="ew", padx=5)
        self.frm_fileAction.pack(padx=10, pady=5)

        #cell change buttons
        self.frm_actions = tk.Frame(self)
        lbl = tk.Label(self.frm_actions, text="Cell:", font=("Courier", 16)).grid(row=0,column=0)
        self.cell = tk.Entry(self.frm_actions, font=("Courier", 16))
        self.cell.grid(row=0,column=1, sticky="we")
        lbl2 = tk.Label(self.frm_actions, text="New value:", font=("Courier", 16)).grid(row=0,column=2)
        self.value = tk.Entry(self.frm_actions, font=("Courier", 16))
        self.value.grid(row=0,column=3)
        self.btn_update = tk.Button(self.frm_actions, text="Update", font=('Coursier',16), command=self.update_cell)
        self.btn_update.grid(row=0, column=4, padx=5)
        self.frm_actions.pack(padx=10,pady=5)

    def open_file(self, event):
        items: List[int] = self.file_list.curselection()
        for item in items:
            fileName = self.file_list.get(item)
            startfile(fileName)

    def add_file(self):
        filepaths = askopenfilenames(
        filetypes=[("Microsoft Excel", "*.xlsx")]
        )
        filepaths = [filepath.replace('/', '\\') for filepath in filepaths]
        for filepath in filepaths:
            if filepath not in self.file_list.get(0,tk.END):
                self.file_list.insert(tk.END, filepath)

    def remove_file(self):
        items: List[int] = self.file_list.curselection()
        for item in items:
            self.file_list.delete(item)

    def back_up(self):
        file_names: List[str] = self.file_list.get(0, tk.END)
        for file_name in file_names:
            copyfile(file_name, file_name+".backup")

    def update_cell(self):
        file_names: List[str] = self.file_list.get(0, tk.END)
        c = self.cell.get().upper()
        v = self.value.get()
        if not c or not v:
            return
        clm = c[0]
        row = c[1:]
        if not (clm.isalpha() and row.isdigit() and int(row) > 0):
            self.cell.configure(fg='red')
            return
        self.cell.configure(fg='black')
        for file_name in file_names:
            workbook = load_workbook(filename=file_name)
            sheet = workbook.active
            sheet[c] = v
            workbook.save(filename=file_name)

if __name__ == "__main__":
    App().mainloop()    
