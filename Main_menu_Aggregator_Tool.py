'''
Created on Aug 7, 2020

@author: cf10387
'''

""" This is the Aggregator Tool
    This module Main_Menu_Aggregator_Tool is the main menu which
    gives the user the pop up menu of actions to be performed. """
from datetime import datetime
import threading
import tkinter as tk
from tkinter import ttk
import New_Process_excel as dd
import logging
import getpass
import os


userID = getpass.getuser()


def initializeLogging():
    """Configures common properties of loggers (at the root logger level)"""

    datetimestr = datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p")
    logfilename = 'C:/Python/TestingOutput/' + userID + '-AggregatorTool As of ' + datetimestr + '.log'

    rootLogger = logging.getLogger()
    hdlr = logging.FileHandler(logfilename)
    formatter = logging.Formatter('%(asctime)s %(levelname)s - %(name)s - %(message)s')
    hdlr.setFormatter(formatter)
    rootLogger.addHandler(hdlr)
    rootLogger.setLevel(logging.DEBUG)


initializeLogging()

logger = logging.getLogger(__name__)
logger.info('File Aggregator Tool starting...')

returnflagstatus = 'Good'



def processFiles(userID, dir_filename, first_row_col_headers_var, last_row_col_headers_var, data_row, sheetname):

    if not first_row_col_headers_var.isdigit():
        tk.messagebox.showerror('First Row of Column Headers = ' + first_row_col_headers_var,
            'Please enter a numerical value for the first row of column headers')
        logger.error('Aggregator Tool Menu - First row of column headers (' + first_row_col_headers_var + ') not numeric - user input error')
        return()

    if not last_row_col_headers_var.isdigit():
        tk.messagebox.showerror('Last Row of Column Headers = ' + last_row_col_headers_var,
            'Please enter a numerical value for the last row of column headers')
        logger.error('Aggregator Tool Menu - Second row of column headers (' + last_row_col_headers_var + ') not numeric - user input error')
        return()

    if not data_row.isdigit():
        tk.messagebox.showerror('First Row of Data = ' + data_row,
            'Please enter a numerical value for the first row of data')
        logger.error('Aggregator Tool Menu - First data row (' + data_row + ') not numeric - user input error')
        return()


    FileISThere = os.path.exists(dir_filename)

    if not FileISThere:
        tk.messagebox.showerror(dir_filename + ' - not found!',
            ' Please enter the correct file path for the output file directory')
        logger.error('Output file path (' + dir_filename + ') was incorrect - user input error')
        return()
    ##call other script as a function

    returnflagstatus = dd.ExcelFileProcess2(userID, first_row_col_headers_var, last_row_col_headers_var, data_row, \
                                            sheetname, dir_filename)

    if (returnflagstatus == 'Empty') :
        tk.messagebox.showinfo('Aggregator Tool Status', 'Select file list is empty. Please select "Process" button and select files again.')
        logger.info('File Aggregator Tool file list is empty')
    elif (returnflagstatus != 'Error') :
        logger.info('File Aggregator Tool finished processing all files')



        ROOT.destroy()


""" threaded progress bar for tkinter gui """

class ProgressBar():
    #def __init__(self, parent, row, column, columnspan):
    def __init__(self, parent):
        self.maximum = 100
        self.interval = 10
        self.progressbar = ttk.Progressbar(parent, orient=tk.HORIZONTAL,
                                           mode="indeterminate",
                                           maximum=self.maximum)
        self.progressbar.place(x=275, y=300, width=250)

        self.thread = threading.Thread()
        self.thread.__init__(target=self.progressbar.start(self.interval),
                             args=())
        self.thread.start()

    def pb_stop(self):
        """ stops the progress bar """
        if not self.thread.isAlive():
            VALUE = self.progressbar["value"]
            self.progressbar.stop()
            self.progressbar["value"] = VALUE

    def pb_start(self):
        """ starts the progress bar """
        if not self.thread.isAlive():
            VALUE = self.progressbar["value"]
            self.progressbar.configure(mode="indeterminate",
                                       maximum=self.maximum,
                                       value=VALUE)
            self.progressbar.start(self.interval)

class mainScreenGUI(tk.Frame):
    def __init__(self,parent,):
        tk.Frame.__init__(self,master=parent)
        #main_screen = tk.Tk()
        parent.title( "File Aggregator Tool v1.0")
        parent.geometry("700x400")
        ttk.Label(parent, text = "  ").grid(row=1, column=3)
        ttk.Label(parent, text = "  ").grid(row=2, column=3)

        #ttk.Label(parent, text = "Enter New, Replace or Append:",
        #          background = 'green', foreground ="white",
        #          font = ('calibri', 10, 'bold')).place(x=10, y=40)

        #create the ComboBox
        #comboval=tk.StringVar()

        #ToolOptionList = ttk.Combobox(parent, width = 15, textvariable = comboval)
        #ToolOptionList['values'] = ["New", "Append","Replace"]

        #ToolOptionList.place(x=275,y=40)
        #ToolOptionList.current(0)

        #action_request = comboval.get()
        first_row_col_headers_var=tk.StringVar()
        last_row_col_headers_var=tk.StringVar()
        first_row_data_var=tk.StringVar()
        tab_name_var=tk.StringVar()
        file_dir_var=tk.StringVar()

        ttk.Label(parent, text = "Enter the first row of column headers:",
                  background = 'green', foreground ="white",
                  font = ('calibri', 10, 'bold')).place(x=10, y=40)

        ttk.Label(parent, text = "Enter last row of column headers:",
                  background = 'green', foreground ="white",
                  font = ('calibri', 10, 'bold')).place(x=10, y=90)

        ttk.Label(parent, text = "Enter first row of data:",
                  background = 'green', foreground ="white",
                  font = ('calibri', 10, 'bold')).place(x=10, y=140)



        tk.Entry(parent, textvariable=first_row_col_headers_var).place(x=275, y=40, width = 40)
        tk.Entry(parent, textvariable=last_row_col_headers_var).place(x=275, y=90, width = 40)
        tk.Entry(parent, textvariable=first_row_data_var).place(x=275, y=140, width = 40)

        #data_row = first_row_data_var.get()

        ttk.Label(parent, text = "Enter Tab name:",
                  background = 'green', foreground ="white",
                  font = ('calibri', 10, 'bold')).place(x=10, y=190)

        tk.Entry(parent, textvariable=tab_name_var).place(x=275, y=190, width=250)

        ttk.Label(parent, text = "Enter Output File Directory(Fully Qualified):",
                  background = 'green', foreground ="white",
                  font = ('calibri', 10, 'bold')).place(x=10, y=240)

        ttk.Label(parent, text = "e.g (//namdfs/ctx/rut/FS_CorpTax_TTI_Data/fs_CRPTX_Federal)",
                font = ('calibri', 10, 'bold')).place(x=25, y=260)

        tk.Entry(parent, textvariable=file_dir_var).place(x=275, y=240, width=350)

        #sheetname = tab_name_var.get()

        ttk.Label(parent, text= "Process Run Status:",
                  background = 'green', foreground = 'white',
                  font = ('calibri', 10, 'bold')).place(x=10, y=300)

        ProgressBar(parent)

        #tk.Button(parent, text='Process', command=lambda: processFiles(comboval.get(), first_row_data_var.get(), \
        #  tab_name_var.get(), file_dir_var.get())).place(x=575, y=250)
        tk.Button(parent, text='Process', command=lambda: processFiles(userID, file_dir_var.get(), first_row_col_headers_var.get(), \
                                                                       last_row_col_headers_var.get(), \
                                                                       first_row_data_var.get(), tab_name_var.get())).place(x=575, y=340)

        tk.Button(parent, text='Quit', command=parent.destroy).place(x=650, y=340)




ROOT=tk.Tk()
APP=mainScreenGUI(ROOT)

ROOT.mainloop()


