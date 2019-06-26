#!./env/Scripts/python
import tkinter as tk
from tkinter import filedialog
import llx2w as x2w
import os

class Application(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        # Configure
        self.title("LongList Generator 0.1.0")
        # self.geometry("350x110")
        self.resizable(0,0)

        self.main = MainWindow(master=self)

        
        # Add NavBar
        self['menu'] = NavBar(master=self)

        # Pack up
        self.main.pack(fill=tk.BOTH, expand=True)

class NavBar(tk.Menu):   
    def __init__(self, master):
        tk.Menu.__init__(self, master=None)
        self.master = master

        self.file = tk.Menu(self)
        self.add_cascade(label="File", menu=self.file)
        self.add_cascade(label="About", menu=self.file)
        self.add_cascade(label="Extra", menu=self.file)
        self.add_cascade(label="Help", menu=self.file)

class Salesboard(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master
        
        OPTIONS = ("Evo", "Nuvi", "NuStar")
        self.var = tk.StringVar(self)

        self.board = tk.Label(self, text="SALESBOARD:")
        self.board.pack(side="left")
        self.dropdown = tk.OptionMenu(self, self.var, *OPTIONS)
        self.dropdown.pack(side="left", fill="x", expand=True)

class UploadFile(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master

        self.PWD = None
        self.filename = tk.StringVar()
        self.topfile = None

        # First Label
        self.upload = tk.Label(self, text="FILENAME:")
        self.upload.pack(side="left")

        self.chosen_file = tk.Label(self, textvariable=self.filename)
        self.chosen_file.pack(side="left", fill=tk.BOTH, expand=True, padx=3)
        self.chosen_file.configure(background="white")

        self.browse_btn = tk.Button(self, text="...", command=self.displayfilename)
        self.browse_btn.pack(side="left")
    
    def displayfilename(self):
        self.PWD = tk.filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*") ))
        self.filename.set(self.PWD.split("/")[-1])
        self.topfile = self.PWD

    def getFilename(self):
        return str(self.filename)

class Example(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master
        
        self.example = tk.Label(self, text="FORMAT:")
        self.example.pack(side="left")
        self.example_name = tk.Label(self, text="Client_PROJECT_N 4 LL for Upload_MM.DD.YY.xlsx")
        self.example_name.pack(side="left")

class Generate(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master = master

        self.generate = tk.Button(self, text="Generate", command=self.genLL)
        self.generate.pack(side="left", fill="x", expand=True)

    def genLL(self):
        path = self.master.uploader.topfile
        if path:
            res = x2w.main(path)
            # self.message = tk.Message(Alert(master = self), text = "Great Job")
            res = "/".join(self.master.uploader.PWD.split("/")[:-1])+"/"+res
            self.message = Alert(master = self, path=res, result=True)
            # self.message.pack(side="left")

        else:
            self.message = Alert(master = self)

class Alert(tk.Toplevel):
    def __init__(self, master, path=None, result=False):
        tk.Toplevel.__init__(self)
        self.frame = tk.Frame(self)
        self.resizable(0,0)
        self.path = path
        self.result = result

        if result:
            self.title("Job Complete")
            self.pathlabel = tk.Label(self.frame, text="Path: {}".format(self.path))
            self.pathlabel.pack(side = "top", fill = tk.X, padx=10, pady=10)
            self.button = tk.Button(self.frame, text="Open Directory", command=lambda:self.open(self.path))
            self.button.pack(side="top", fill=tk.X, anchor = "center", padx = 10, pady=5)
        else:
            self.title("Job Failure")
            self.geometry("250x50")
            self.message = tk.Label(self.frame, text="You need to select a file")
            self.message.place(relx=.5, rely=.5, anchor="center")
        
        self.frame.pack(fill=tk.BOTH, expand=True)

    def open(self, f):
        folder = "/".join(f.split("/")[:-1])
        folder = os.path.realpath(folder)
        os.startfile(folder)


class MainWindow(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self)

        self.master = master
        # self.salesboard = Salesboard(master=self)
        # self.salesboard.pack(side="top", fill="x", padx=10, pady=5) #side=tk.LEFT,

        self.uploader = UploadFile(master=self)
        self.uploader.pack(side="top", fill="x", padx=10, pady=5)

        self.example = Example(master=self)
        self.example.pack(side="top", fill="x", padx=10, pady=5)

        self.generate_btn = Generate(master=self)
        self.generate_btn.pack(side="top", fill="x", padx=10, pady=5)
        # self.salesboard.configure(background="black")

      
if __name__ == "__main__":
    app = Application()
    app.mainloop()
