import tkinter as tk 
import os
import checker.checker as c
from tkinter import filedialog
import configparser 

settings = c.readSettings("checker/settings.ini")

class setVariables():
    def __init__(self, master):
        self.mainexcel = tk.StringVar()
        self.template = tk.StringVar()
        self.uob_t = tk.StringVar()
        self.mainexcelpath = ""
        self.templatepath = ""
        self.uob_tpath = ""
        self.master = master

        # file selection frame ------------------------------------------------------------------------------
        self.fileframe = tk.LabelFrame(self.master, text="Setup file paths", padx=5, pady=5)
        self.fileframe.grid(sticky = "W",row=0, column=0)

        # main excel file
        self.path_label = tk.Label(self.fileframe, text = "Location of file to check: ", anchor="w")
        self.path = tk.Entry(self.fileframe, width=30, textvariable=self.mainexcel)
        self.path.insert(0, settings["path"])
        self.selectexcelbutton = tk.Button(self.fileframe, text="Select", command=self.selectFile)

        self.path_label.grid(sticky = "W",row=3, column=0)
        self.path.grid(row=3, column=1)
        self.selectexcelbutton.grid(row=3, column=2)

        # file to populate
        self.to_check_label = tk.Label(self.fileframe, text = "Location to populate: ", anchor="w")
        self.to_check = tk.Entry(self.fileframe, width=30, textvariable=self.template)
        self.to_check.insert(0, settings["template"])
        self.selecttemplatebutton = tk.Button(self.fileframe, text="Select", command=self.selectTemplate)

        self.to_check_label.grid(sticky = "W",row=4, column=0)
        self.to_check.grid(row=4, column=1)
        self.selecttemplatebutton.grid(row=4, column=2)

        # UOB lookup file
        self.uob_label = tk.Label(self.fileframe, text = "Location of UOB lookup file: ", anchor="w")
        self.uob = tk.Entry(self.fileframe, width=30, textvariable=self.uob_t)
        self.uob.insert(0, settings["uob"])
        self.selectuobbutton = tk.Button(self.fileframe, text="Select", command=self.selectUOB)

        self.uob_label.grid(sticky = "W",row=5, column=0)
        self.uob.grid(row=5, column=1)
        self.selectuobbutton.grid(row=5, column=2)

        # checker button frame --------------------------------------------------------------------------
        self.checkerframe = tk.LabelFrame(self.master, text="Click to run", padx=5, pady=5)
        self.checkerframe.grid(row=10, column=10)

        # checker button
        self.checkbutton = tk.Button(self.checkerframe, text="Check", command=self.runCheck, width=10)
        self.checkbutton.grid(row=10, column=5)

        # input frames --------------------------------------------------------------------------
        self.inputframe = tk.LabelFrame(self.master, text = "Fields", padx=5, pady=5)
        self.inputframe.grid(sticky = "W",row=7, column=0)

        # inputs --------------------------------------------------------------------------------

        # date
        self.date = tk.Label(self.inputframe, text="Date: ", anchor="w")
        self.d = tk.Entry(self.inputframe, width=10)
        self.d.insert(0, settings["date"])

        self.date.grid(row=8, column=0)
        self.d.grid(row=8, column=1)

        # fiscal period
        self.fiscalperiod = tk.Label(self.inputframe, text="Fiscal period: ", anchor="w")
        self.fiscal = tk.Entry(self.inputframe, width=10)
        self.fiscal.insert(0, settings["fp"])

        self.fiscalperiod.grid(row=9, column=0)
        self.fiscal.grid(row=9, column=1)

        # Document
        self.assignmentno = tk.Label(self.inputframe, text="Document: ", anchor="w")
        self.assignment = tk.Entry(self.inputframe, width=10)
        self.assignment.insert(0, settings["assignmentnumber"])

        self.assignmentno.grid(row=10, column=0)
        self.assignment.grid(row=10, column=1)

        # save settings ---------------------------------------------------------------------------
        self.savesettings = tk.Button(self.checkerframe, text="Save", command=self.saveSettings, width=10)
        self.savesettings.grid(row=9, column=5)

        # Populate Button
        self.populatebutton = tk.Button(self.checkerframe, text="Populate", command=self.fillTemplate, width=10)
        self.populatebutton.grid(row=11, column=5)   
        
    def fillTemplate(self):
        os.system("python checker\\populate.pyw")

    def runCheck(self):
        self.check = tk.Label(self.master, text = "Check complete")
        os.system('python checker\\checker.pyw')
    
    def selectFile(self):
        self.mainexcelpath = tk.filedialog.askopenfilename(initialdir=os.path.dirname(settings["path"]), title="Select file", filetypes=(("Excel files", ".xlsx .xls"), ("All files", "*.*")))
        self.mainexcel.set(self.mainexcelpath)

    def selectTemplate(self):
        self.templatepath = (filedialog.askopenfilename(initialdir=os.path.dirname(settings["template"]), title="Select file", filetypes=(("Excel files", "*.xlsx .xls"), ("All files", "*.*"))))
        self.template.set(self.templatepath)

    def selectUOB(self):
        self.uob_tpath = (filedialog.askopenfilename(initialdir=os.path.dirname(settings["uob"]), title="Select file", filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))))
        self.uob_t.set(self.uob_tpath)
        
    def saveSettings(self):
        self.config = configparser.ConfigParser()
        self.config.read("checker\\settings.ini")

        self.config["MAIN"]["assignmentnumber"] = self.assignment.get()
        self.config["MAIN"]["fp"] = self.fiscal.get()
        self.config["MAIN"]["uob"] = self.uob.get()
        self.config["MAIN"]["template"] = self.to_check.get()
        self.config["MAIN"]["path"] = self.path.get()
        self.config["MAIN"]["date"] = self.d.get()

        with open("checker/settings.ini", "w") as f:
            self.config.write(f)

        self.newWindow = tk.Toplevel(self.master)
        self.app = Saved(self.newWindow)


class Saved:
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.text = tk.Label(self.frame, text = "Settings saved.")
        self.okbutton = tk.Button(self.frame, text="Ok", width=10, command = self.close_window)
        self.text.pack()
        self.okbutton.pack()
        self.frame.pack()

    def close_window(self):
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    root.title('Check and Populate')
    app = setVariables(root)
    root.mainloop()
