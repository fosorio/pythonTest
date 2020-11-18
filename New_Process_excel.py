


'''
Created on Oct 15, 2020

@author: cf10387
'''

import jpype
import asposecells
from importlib_metadata import files
jpype.startJVM()
from asposecells.api import *
from tkinter import filedialog
from datetime import datetime
from tkinter import messagebox
import logging
import os

NORM_FONT = ("Calibri", 10)
MAX = 30

logger = logging.getLogger(__name__)

def ExcelFileProcess2(userID, first_row_col_headers, last_row_col_headers, data_row, strsheetname, dir_filename):

    excel_files_list=filedialog.askopenfilenames(title='Select files to be opened')

    if not excel_files_list :
        returnflag = 'Empty'
        return(returnflag)

    i = 0
    returnflag = 'Good'
    firstFileCtr = 0
    ProcessLogCtr = 1
    #mylist = list(range(HeaderRowTotal))
    #print(mylist)

    logger.info('Reading Excel files in 2nd module')

    for files in excel_files_list:

        if firstFileCtr == 0 :

            #create new workbook for aggregated files
            datetimestr = datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p")

            wbk = Workbook(files)
            wbk2 = Workbook()

            worksheetCtr = wbk.getWorksheets().getCount()
            firstFileCtr = firstFileCtr + 1
            for i in range(worksheetCtr) :
                        #wksht = wbk.getWorksheets().get(strsheetname)
                wkshtSource = wbk.getWorksheets().get(i)
                newworkbooksheetname = wkshtSource.getName()
                        #if wbk.getWorksheets().get(i) == strsheetname :
                if newworkbooksheetname == strsheetname :
                    i=1000
                    continue
                else  :
                    messagebox.showerror('Sheet Name = ' + strsheetname,
                                          'Please enter the correct tab_name on the third question')
                    logger.error(files + ' - Tab name ' + strsheetname + ' was incorrect - user input error')
                    returnflag = 'Error'
                    return(returnflag)
                i = i + 1

            wksht2Target = wbk2.getWorksheets().add("Aggregator Tool")
            wksht3ProgressLog = wbk2.getWorksheets().add("Process Log")

            ProcessLogHeader = wksht3ProgressLog.getCells().get("A1")
            ProcessLogHeader.putValue("Source File Name and Path")
            ProcessLogHeader = wksht3ProgressLog.getCells().get("B1")
            ProcessLogHeader.putValue("Process Date and Time")
            ProcessLogHeader = wksht3ProgressLog.getCells().get("C1")
            ProcessLogHeader.putValue("Messages")
            ProcessLogHeader = wksht3ProgressLog.getCells().get("D1")
            ProcessLogHeader.putValue("Aggregated Data Tab Row Count of Data Added")

            ProcessLogCtr = ProcessLogCtr + 1

            thecell = wksht3ProgressLog.getCells().get("A1")
            style = thecell.getStyle()
            style.getFont().setName("Calibri")
            style.getFont().setSize(12)
            style.getFont().setUnderline(FontUnderlineType.SINGLE)
            style.getFont().setBold(True)
            thecell.setStyle(style)

            thecell = wksht3ProgressLog.getCells().get("B1")
            style = thecell.getStyle()
            style.getFont().setName("Calibri")
            style.getFont().setSize(12)
            style.getFont().setUnderline(FontUnderlineType.SINGLE)
            style.getFont().setBold(True)
            thecell.setStyle(style)

            thecell = wksht3ProgressLog.getCells().get("C1")
            style = thecell.getStyle()
            style.getFont().setName("Calibri")
            style.getFont().setSize(12)
            style.getFont().setUnderline(FontUnderlineType.SINGLE)
            style.getFont().setBold(True)
            thecell.setStyle(style)

            thecell = wksht3ProgressLog.getCells().get("D1")
            style = thecell.getStyle()
            style.getFont().setName("Calibri")
            style.getFont().setSize(12)
            style.getFont().setUnderline(FontUnderlineType.SINGLE)
            style.getFont().setBold(True)
            style.setTextWrapped(True)
            thecell.setStyle(style)



            wksht3ProgressLog.getCells().setColumnWidth(0, 60.0 )
            wksht3ProgressLog.getCells().setColumnWidth(1, 35.0 )
            wksht3ProgressLog.getCells().setColumnWidth(2, 75.0 )
            wksht3ProgressLog.getCells().setColumnWidth(3, 25.0 )

            # style.setBorder()
            #wksht3.setStyle(style)

            lastRowSourceFile = wkshtSource.getCells().getMaxDataRow()

            if int(lastRowSourceFile) == 0 or int(lastRowSourceFile) < int(data_row) :
                messagetext='Error - This file contains no rows of data from the first data row value that was selected. This file was not processed.'
                datetimestr = datetime.now()

                rowcount = 0
                thecell = wksht3ProgressLog.getCells().get("A" + ProcessLogCtr)
                thecell.putValue = files
                thecell = wksht3ProgressLog.getCells().get("B" + ProcessLogCtr)
                thecell.putValue = datetimestr
                thecell = wksht3ProgressLog.getCells().get("C" + ProcessLogCtr)
                thecell.putValue = messagetext
                thecell = wksht3ProgressLog.getCells().get("D" + ProcessLogCtr)
                thecell.putValue = rowcount

                ProcessLogCtr = ProcessLogCtr + 1

                logger.error(files + ' - file contains no rows of data')
                continue
            else :
                lastColSourceFile = wkshtSource.getCells().getMaxDataColumn()

                wksht2Target.copy (wbk.getWorksheets().get(0))

                wksht2Target.getCells().insertColumns(0,2)
                #wksht2Target.getCells().get("A" + str(offsetHeaderRow)).setValue("Filename")

                #therangeTarget = wksht2Target.getCells().createRange("A" + str(data_row), "A" + str(lastRowSourceFile + 1))
                strlastRowSourceFile = lastRowSourceFile + 1
                therangeTarget = wksht2Target.getCells().createRange("A1", "A" + str(strlastRowSourceFile))

                therangeTarget.setValue(files)
                #wksht2Target.getCells().get("A" + str(offsetHeaderRow)).setValue("Filename")

                wksht2Target.getCells().setColumnWidth(0, 60.0 )
                adjHeaderrowrange = int(first_row_col_headers) - 1

                #wksht2Target.getCells().insertColumns(0,2)
                theheaderrangeTarget = wksht2Target.getCells().createRange("B1","B" + str(adjHeaderrowrange))

                theheaderrangeTarget.setValue("Headers")

                thecolheadersrangeTarget = wksht2Target.getCells().createRange("B" + str(first_row_col_headers), \
                                                                           "B" + str(last_row_col_headers))

                thecolheadersrangeTarget.setValue("Column Headers")

                thedatarangeTarget = wksht2Target.getCells().createRange("B" + str(data_row), \
                                                                           "B" + str(strlastRowSourceFile))

                thedatarangeTarget.setValue("Data")

                rowcount = int(lastRowSourceFile) + 1

                messagetext='File processed. Data added.'
                datetimestr = datetime.now()


                thecell = wksht3ProgressLog.getCells().get("A" + str(ProcessLogCtr))
                thecell.setValue(files)
                thecell = wksht3ProgressLog.getCells().get("B" + str(ProcessLogCtr))
                thecell.setValue(str(datetimestr))
                thecell = wksht3ProgressLog.getCells().get("C" + str(ProcessLogCtr))
                thecell.setValue(messagetext)
                thecell = wksht3ProgressLog.getCells().get("D" + str(ProcessLogCtr))
                thecell.setValue(rowcount)

                ProcessLogCtr = ProcessLogCtr + 1

                success_string = (files + ' - ' + str(datetimestr) + ' - ' + messagetext + ' - ' + str(rowcount) + ' row(s)')
                logger.info(success_string)

        else :
            wbk = Workbook(files)

            worksheetCtr = wbk.getWorksheets().getCount()

            for i in range(worksheetCtr) :
                #wksht = wbk.getWorksheets().get(strsheetname)
                wkshtSource = wbk.getWorksheets().get(i)
                newworkbooksheetname = wkshtSource.getName()
                #if wbk.getWorksheets().get(i) == strsheetname :
                if newworkbooksheetname == strsheetname :
                    i=1000
                    continue
                else  :
                    messagetext='Error - This file does not contain the specified worksheet tab name. This file was not processed.'
                    logger.error(files + ' - Tab name ' + strsheetname + ' was incorrect - user input error')
                    datetimestr = datetime.now()
                    rowcount = 0


            lastRowTargetFile = wksht2Target.getCells().getMaxDataRow()
            lastColTargetFile = wksht2Target.getCells().getMaxDataColumn()
            lastRowSourceFile = wkshtSource.getCells().getMaxDataRow()
            lastColSourceFile = wkshtSource.getCells().getMaxDataColumn()

            #lastCell = wkshtSource.getCells().getLastCell()

            if int(lastRowSourceFile) == 0 or int(lastRowSourceFile) < int(data_row) :
                messagetext='Error - This file contains no rows of data from the first data row value that was selected. This file was not processed.'
                datetimestr = datetime.now()

                rowcount = 0
                thecell = wksht3ProgressLog.getCells().get("A" + str(ProcessLogCtr))
                thecell.setValue(files)
                thecell = wksht3ProgressLog.getCells().get("B" + str(ProcessLogCtr))
                thecell.setValue(str(datetimestr))
                thecell = wksht3ProgressLog.getCells().get("C" + str(ProcessLogCtr))
                thecell.setValue(messagetext)
                thecell = wksht3ProgressLog.getCells().get("D" + str(ProcessLogCtr))
                thecell.setValue(rowcount)

                ProcessLogCtr = ProcessLogCtr + 1

                logger.error(files + ' file contains no rows of data')
                continue
            else :
                adjlastRowSourceFile = lastRowSourceFile + 1
                adjlastColSourceFile = lastColSourceFile + 1
                # rowcount = int(lastRowSourceFile) - int(data_row) + 1
                rowcount = int(lastRowSourceFile) + 1
                rangeSource = wkshtSource.getCells().createRange(0, 0, adjlastRowSourceFile, \
                                                                    adjlastColSourceFile)
                rangeTarget = wksht2Target.getCells().createRange(lastRowTargetFile + 1, 2, \
                                                                     adjlastRowSourceFile, adjlastColSourceFile)
                rangeTarget.copy(rangeSource)

                strlastRowSourceFile = lastRowSourceFile + 1

                #therangeTarget = wksht2Target.getCells().createRange("A1", "A" + str(strlastRowSourceFile))
                therangeFilename = wksht2Target.getCells().createRange("A" + str(lastRowTargetFile + 2) + \
                                                                     ":A" + str(lastRowTargetFile + 1 + rowcount) )
                therangeFilename.setValue(files)

                adjHeaderrowrange = int(lastRowTargetFile) + 2 + int(first_row_col_headers) - 2
                theheaderrangeTarget = wksht2Target.getCells().createRange("B" + str(adjHeaderrowrange), \
                                                                            "B" + str(int(adjHeaderrowrange) + int(first_row_col_headers) - 2))

                theheaderrangeTarget.setValue("Headers")


                adjcolHeaderrowrange = int(lastRowTargetFile) + 2 + int(first_row_col_headers) - 1
                bottomcolHeaderrowrange = int(lastRowTargetFile) + 2 + int(last_row_col_headers) - 1

                thecolheaderrangeTarget = wksht2Target.getCells().createRange("B" + str(adjcolHeaderrowrange) + ":B" + str(bottomcolHeaderrowrange))

                thecolheaderrangeTarget.setValue("Column Headers")

                adjdatarowstartrange = int(bottomcolHeaderrowrange) + 1
                adjdatarowendrange = int(lastRowTargetFile) + 1 + int(rowcount)

                thedatarangeTarget = wksht2Target.getCells().createRange("B" + str(adjdatarowstartrange) + \
                                                                     ":B" + str(adjdatarowendrange))

                thedatarangeTarget.setValue("Data")

                messagetext='File processed. Data added.'
                datetimestr = datetime.now()


                thecell = wksht3ProgressLog.getCells().get("A" + str(ProcessLogCtr))
                thecell.setValue(files)
                thecell = wksht3ProgressLog.getCells().get("B" + str(ProcessLogCtr))
                thecell.setValue(str(datetimestr))
                thecell = wksht3ProgressLog.getCells().get("C" + str(ProcessLogCtr))
                thecell.setValue(messagetext)
                thecell = wksht3ProgressLog.getCells().get("D" + str(ProcessLogCtr))
                thecell.setValue(rowcount)

                ProcessLogCtr = ProcessLogCtr + 1

                success_string = (files + ' - ' + str(datetimestr) + ' - ' + messagetext + ' - ' + str(rowcount) + ' row(s)')
                logger.info(success_string)

                    #datetimestr = datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p")
    logger.info('Finished reading Excel files in 2nd module')
    outfilename = dir_filename + '/' + userID + '-Aggregator_Solution_AsOf_' + datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p") + '.xlsb'


    wbk2.save(outfilename)


    return