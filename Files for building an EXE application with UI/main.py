from tkinter import *
from tkinter.ttk import *
import sys
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog, messagebox, _setit
import os, warnings, pandas as pd
import threading
import multiprocessing
from queue import Empty
from PIL import Image, ImageTk
import webbrowser
from opendatabim.report import Report
from datetime import datetime, timedelta
import tkinter.font as tkFont

warnings.simplefilter(action='ignore', category=UserWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=pd.errors.DtypeWarning)

class PrintLogger(object):

    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, text):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", text)
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def flush(self):
        pass


class ReportAPP(Tk):
    def __init__(self):

        Tk.__init__(self)

        try:
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")

        icon = os.path.join(self.base_path, 'img', 'icon.png')
        icon = PhotoImage(file=icon)
        self.call('wm', 'iconphoto', self._w, icon)

        theme = os.path.join(self.base_path, "theme", "sv.tcl")
        self.call("source", theme)
        self.call("set_theme", "light")


        self.title('OpenDataBIM')

        self.setupGUI()
        self.setStyles()

    def setupGUI(self):

        self.setGeometry()

        self.resizable(False, False)

        mainFrame = Frame(self)
        mainFrame.pack(side=LEFT, fill=BOTH, expand=1)

        frameL = Frame(mainFrame)
        frameR = Frame(mainFrame)

        frameL.grid(row=0, column=0, padx=10, pady=10, ipadx=0, ipady=0, sticky="NW")
        frameR.grid(row=0, column=1, padx=10, pady=8, ipadx=0, ipady=0, sticky="NE")

        header = Label(frameL)

        header.grid(row=0, column=0, padx=10, sticky="NW")
        headerText = Label(header)
        headerText.grid(row=0, column=0, padx=10, sticky="W")


        Label(headerText,
              text="Checking BIM models",
              style="Header1.TLabel"
              ).grid(
                    row=0,
                    column=0,
                    sticky="W"
                    )

        Label(headerText,
              text="Revit and IFC projects",
              style="Header2.TLabel"
              ).grid(
                    row=1,
                    column=0,
                    sticky="W"
                    )


        Link = Label(headerText,
                    text="Read more about the check too",
                    style="Link.TLabel",
                    cursor="hand2"
                     )
        Link.grid(
                    row=2,
                    column=0,
                    sticky="W"
                )
        Link.bind("<Button-1>",
                lambda _: webbrowser.open_new("https://opendatabim.io/index.php/quality-of-revit-and-ifc-projects/")
                  )
        odbLogo = os.path.join(self.base_path, 'img', 'odb.png')
        odbLogo = Image.open(odbLogo)
        odbLogo = odbLogo.resize((int(612*0.19), int(422*0.19)))

        odbLogo = ImageTk.PhotoImage(odbLogo)

        headerLogo = Label(header, image=odbLogo, cursor="hand2")
        headerLogo.image = odbLogo
        headerLogo.grid(
                    row=0,
                    column=1,
                    sticky="W")

        headerLogo.bind("<Button-1>", lambda _: webbrowser.open_new("https://opendatabim.io/"))


        Label(frameL, text="📁 Path to folder with Revit or IFC files", style="Option.TLabel").grid(
                    row=1,
                    column=0,
                    padx=10,
                    pady=(7,0),
                    sticky="W",

                )

        pathToFolder = Entry(frameL,style="Placeholder.TEntry")
        pathToFolder.grid(
                    row=2,
                    column=0,
                    padx=10,
                    sticky="WE"
                )

        Button(frameL, text="Select Folder", width=15, command=self.get_pathToFolder, style="Button.TButton").grid(
                    row=2,
                    column=1,
                    sticky="W"
                )


        Label(frameL, text="🗃 Check all the projects in the subfolders?", style="Option.TLabel").grid(
                    row=3,
                    column=0,
                    padx=10,
                    pady=(7, 0),
                    sticky="W",

        )
        varCheckSubfolders = StringVar()
        varCheckSubfolders.set("No")
        checkSubfolders = OptionMenu(frameL, varCheckSubfolders, "No", "Yes", "No",  style="Dropdown.TMenubutton")
        checkSubfolders.grid(
                    row=4,
                    column=0,
                    padx=10,
                    sticky="WE"
        )





        Label(frameL, text="📁 Results output folder", style="Option.TLabel").grid(
                    row=5,
                    column=0,
                    padx=10,
                    pady=(10, 2),
                    sticky="W"
                )
        pathOutputFolder = Entry(frameL, style="Placeholder.TEntry")
        pathOutputFolder.grid(
                    row=6,
                    column=0,
                    padx=10,
                    pady=2,
                    sticky="WE"
                )
        Button(frameL, text="Select Folder", width=15, style="Button.TButton", command=self.get_pathOutputFolder).grid(
                    row=6,
                    column=1,
                    sticky="W"
                )

        Label(frameL, text="📁 Path to noBIM converter files folder", style="Option.TLabel").grid(
                    row=7,
                    column=0,
                    padx=10,
                    pady=(10, 2),
                    sticky="W"
                )

        pathConverterFolder = Entry(frameL, style="Placeholder.TEntry")
        pathConverterFolder.grid(
                    row=8,
                    column=0,
                    padx=10,
                    sticky="WE"
                )

        Button(frameL, text="Select Folder", width=15, style="Button.TButton", command=self.get_pathConverterFolder).grid(
                    row=8,
                    column=1,
                    sticky="W"
                )

        imgConverter = os.path.join(self.base_path, 'img', 'converter.png')
        imgConverter = Image.open(imgConverter)
        imgConverter = imgConverter .resize((int(2034*0.25), int(687*0.25)))
        imgConverter = ImageTk.PhotoImage(imgConverter)
        placeImgConverter = Label(frameL, image=imgConverter, cursor="hand2")
        placeImgConverter.image = imgConverter
        placeImgConverter.grid(
                    row=9,
                    column=0,
                    columnspan=2,
                    sticky="W",
                    padx=10,
                    pady=(10, 0),
        )

        placeImgConverter.bind("<Button-1>", lambda _: webbrowser.open_new("https://opendatabim.io/"))

        Label(frameR, text="📑 Excel spreadsheet with checking parameters", style="Option.TLabel").grid(
                    row=1,
                    column=0,
                    padx=10,
                    pady=(0, 2),
                    sticky="W"
                )

        pathFileExcel = Entry(frameR, style="Placeholder.TEntry")
        pathFileExcel.grid(
                    row=2,
                    column=0,
                    padx=10,
                    sticky="WE"
                )

        Button(frameR, text="Select File", width=15, command=self.get_pathFileExcel, style="Button.TButton").grid(
                    row=2,
                    column=1,
                    sticky="W"
                )

        Label(frameR, text="📁 Path to PDF_Sources folder", style="Option.TLabel").grid(
                    row=3,
                    column=0,
                    padx=10,
                    pady=(10, 2),
                    sticky="W"
                )
        pathPdfSources = Entry(frameR, font=tkFont.Font(family='Helvetica', size=36, weight='bold'))
        pathPdfSources.grid(
                    row=4,
                    column=0,
                    padx=10,
                    pady=2,
                    sticky="WE"
                )
        Button(frameR, text="Select Folder", width=15, command=self.get_pathPdfSources, style="Button.TButton").grid(
                    row=4,
                    column=1,
                    sticky="W"
                )

        Label(frameR, text="📘 What projects are checked?", style="Option.TLabel", width=55).grid(
                    row=5,
                    column=0,
                    padx=10,
                    pady=(10, 2),
                    sticky="W"
                )
        varProjectType = StringVar()
        varProjectType.set("Revit")
        projectType = OptionMenu(frameR, varProjectType, "Revit", "Revit", "IFC", style="Dropdown.TMenubutton", command=self.set_groupingParam)
        projectType.grid(
                    row=6,
                    column=0,
                    padx=10,
                    sticky="WE"
        )

        Label(frameR, text="Grouping Parameter", style="Option.TLabel", width=55).grid(
                    row=7,
                    column=0,
                    padx=10,
                    pady=(10, 2),
                    sticky="W"
                )

        varGroupingParam = StringVar()
        varGroupingParam.set("Category")
        groupingParam = OptionMenu(frameR, varGroupingParam, "Category", "Category", "Type Name", style="Dropdown.TMenubutton")
        groupingParam.grid(
                    row=8,
                    column=0,
                    padx=10,
                    sticky="WE"
        )

        self.next_check = None
        Label(frameR, text="⏱ Run a check at the same time every 24 hours?", style="Option.TLabel", width=55).grid(
                row=9,
                column=0,
                padx=10,
                pady=(10, 2),
                sticky="W"
            )

        varCheckEvery24Hours = StringVar()
        varCheckEvery24Hours.set("No")
        checkEvery24Hours = OptionMenu(frameR, varCheckEvery24Hours, "No", "Yes", "No", style="Dropdown.TMenubutton")
        checkEvery24Hours.grid(
                    row=10,
                    column=0,
                    padx=10,
                    sticky="WE"
        )

        progressBar = Progressbar(frameR, orient=HORIZONTAL, mode='indeterminate', length=280)
        progressBar.grid(
                    row=11,
                    column=0,
                    padx=10,
                    pady=(10,0),
                    sticky="WE"
                )

        infoPlace = Label(frameR)
        infoPlace.grid(
                row=12,
                column=0,
                columnspan=2,
                padx=10,
                pady=(25, 0),
                sticky="W"
                )


        imgODB = os.path.join(self.base_path, 'img', 'logo.png')

        imgODB = Image.open(imgODB)

        imgODB = ImageTk.PhotoImage(imgODB.resize((int(555*0.4), int(125*0.4))))

        placeImgODB = Label(infoPlace, image=imgODB, cursor="hand2")
        placeImgODB.image = imgODB
        placeImgODB.grid(
                    row=1,
                    column=0,
                    sticky="W"
        )

        placeImgODB.bind("<Button-1>", lambda _: webbrowser.open_new("https://opendatabim.io/"))

        infoText = Label(infoPlace, text="Open application code\nand parameterizable output", style="Info.TLabel", cursor = "hand2")
        infoText.grid(
                    row=1,
                    column=1,
                    sticky="W",
                    padx=(10, 0)
                )
        infoText.bind("<Button-1>", lambda _: webbrowser.open_new("https://github.com/OpenDataBIM/Checking-the-quality-of-Revit-and-IFC-projects"))

        buttonStart = Button(frameR, text="Start", width=15, command=lambda: self.startWork("button"), style="Button.TButton")
        buttonStart.grid(
                row=12,
                column=1,
                pady=(25,0),
                sticky="W"
                )
        frameLog = Label(frameR)
        frameLog.grid(row=13, column=0, padx=10, sticky="W")
        widgetLog = ScrolledText(frameLog, width=70, height=7, font=("Consolas", "10", "normal"), )
        widgetLog.grid(column=0,
                             row=0,
                             sticky='W',
                             padx=10,
                             pady=10,
                             )
        widgetLog.config(state=DISABLED)
        logger = PrintLogger(widgetLog)
        sys.stdout = logger
        sys.stderr = logger

        helpFrame = Frame(frameR)
        helpFrame.grid(row=15, column=0, padx=10, sticky="W")
        email = Label(helpFrame, text = "📧 Please send an email to info@opendatabim.io if you come across\nany issues or errors", cursor="hand2", style="About.TLabel")
        email.grid(
                    row=0,
                    column=0,
                    columnspan=2,
                    sticky="W",
                    pady=(5, 30)
                )

        email.bind("<Button-1>",
                lambda _: webbrowser.open_new("mailto:info@opendatabim.io")
                  )

        self.pathToFolder = pathToFolder
        self.pathOutputFolder = pathOutputFolder
        self.pathConverterFolder = pathConverterFolder
        self.pathFileExcel = pathFileExcel
        self.pathPdfSources = pathPdfSources
        self.varProjectType = varProjectType
        self.varGroupingParam = varGroupingParam
        self.varCheckEvery24Hours = varCheckEvery24Hours
        self.varCheckSubfolders = varCheckSubfolders
        self.groupingParam = groupingParam
        self.progressBar = progressBar
        self.buttonStart = buttonStart

        return

    def set_groupingParam(self, e):
        if e == "Revit":
            self.varGroupingParam.set('Category')
            self.groupingParam['menu'].delete(0, 'end')

            for choice in ('Category', 'Type Name'):
                self.groupingParam['menu'].add_command(label=choice, command=_setit(self.varGroupingParam, choice))

        elif e == "IFC":
            self.varGroupingParam.set('IfcEntity')
            self.groupingParam['menu'].delete(0, 'end')

            for choice in ('IfcEntity', 'ObjectType'):
                self.groupingParam['menu'].add_command(label=choice, command=_setit(self.varGroupingParam, choice))

        return

    def get_pathToFolder(self):
        path = filedialog.askdirectory(
            initialdir=self.pathToFolder.get()
        )
        if path:
            path = os.path.normpath(path)
            self.pathToFolder.delete(0, END)
            self.pathToFolder.insert(0, path)

            if not self.pathPdfSources.get():
                self.pathPdfSources.insert(0, os.path.join(path, 'PDF_Sources'))
            if not self.pathOutputFolder.get():
                self.pathOutputFolder.insert(0, path)

        return

    def get_pathPdfSources(self):
        path = filedialog.askdirectory(initialdir=self.pathPdfSources.get())
        if path:
            path = os.path.normpath(path)
            self.pathPdfSources.delete(0, END)
            self.pathPdfSources.insert(0, path)

        return

    def get_pathFileExcel(self):
        path = filedialog.askopenfilename(defaultextension='.xlsx',
                                          filetypes=[("xlsx", "*.xlsx"), ("All files", "*.*")])
        if path:
            path = os.path.normpath(path)
            self.pathFileExcel.delete(0, END)
            self.pathFileExcel.insert(0, path)

        return

    def get_pathOutputFolder(self):
        path = filedialog.askdirectory(initialdir=self.pathOutputFolder.get())
        if path:
            path = os.path.normpath(path)
            self.pathOutputFolder.delete(0, END)
            self.pathOutputFolder.insert(0, path)

        return

    def get_pathConverterFolder(self):
        path = filedialog.askdirectory(initialdir=self.pathConverterFolder.get())
        if path:
            path = os.path.normpath(path)
            self.pathConverterFolder.delete(0, END)
            self.pathConverterFolder.insert(0, path)
        return

    def setStyles(self):

        style = Style()


        style.configure('Header1.TLabel', font=("Poppins", 24, "bold"), foreground="black")
        style.configure('Header2.TLabel', font=("Poppins", 12, "normal"))
        style.configure('Dropdown.TMenubutton', font=("Poppins", 10, "normal"), anchor = 'w')
        style.configure('Option.TLabel', font=("Poppins", 10, "bold"), padding=(0,3,0,0))
        style.configure('Placeholder.TEntry', background="gray", font=("Poppins", 20, "normal"))
        style.configure('Button.TButton', font=("Poppins", 10, "normal"))
        style.configure('Link.TLabel', font=("Poppins", 10, "normal", "underline"), foreground="blue")
        style.configure('Info.TLabel', font=("Poppins", 11, "bold"))
        style.configure('TProgressbar')
        style.configure('About.TLabel', font=("Poppins", 10, "normal"))

    def getBestGeometry(self):

        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()


        self.w = w = ws / 1.07
        self.h = h = hs / 1.55
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)
        g = '%dx%d+%d+%d' % (w, h, x, y)

        return g

    def setGeometry(self):
        self.winsize = self.getBestGeometry()
        self.geometry(self.winsize)
        return

    def startWork(self, initiator):
        if not all([
            self.pathToFolder.get(),
            self.pathOutputFolder.get(),
            self.pathConverterFolder.get(),
            self.pathFileExcel.get(),
            self.pathPdfSources.get(),
        ]):
            messagebox.showerror(message="Please, select all parameters!")
            return

        if self.next_check and initiator == "button":
            print("The previous check schedule has been canceled!")
            self.after_cancel(self.next_check)
            self.next_check = None

        if self.varCheckEvery24Hours.get() == "Yes":
            now = datetime.now()
            next_check_time = now + timedelta(days=1)
            print(next_check_time.strftime("Next check at %Y-%m-%d %H:%M:%S"))
            self.next_check = self.after(24*60*60*1000, lambda: self.startWork("schedule"))

        self.buttonStart["state"] = DISABLED
        self.Thread_ProgressBar = threading.Thread()
        self.Thread_ProgressBar.__init__(target=self.progressBar.start())
        self.Thread_ProgressBar.start()

        data = dict(
            pathToFolder=self.pathToFolder.get(),
            pathOutputFolder=self.pathOutputFolder.get(),
            pathConverterFolder=self.pathConverterFolder.get(),
            pathFileExcel=self.pathFileExcel.get(),
            pathPdfSources=self.pathPdfSources.get(),
            varProjectType=self.varProjectType.get(),
            varGroupingParam=self.varGroupingParam.get(),
            varCheckEvery24Hours=self.varCheckEvery24Hours.get(),
            varCheckSubfolders=self.varCheckSubfolders.get(),
        )

        self.queue = multiprocessing.Queue()
        proc = multiprocessing.Process(target=Report, args=(self.queue, data), daemon=True)
        proc.start()
        self.check()


        return

    def check(self):
        try:
            data = self.queue.get(block=False)
        except Empty:
            pass

        else:
            if data == 'done':
                self.progressBar.stop()
                self.buttonStart["state"] = NORMAL
                return

            print(data)

        finally:
            self.after(1000, self.check)

def main():
    app = ReportAPP()

    app.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
