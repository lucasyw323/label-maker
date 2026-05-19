from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import pandas as pd
import numpy as np
import xlwings as xw
from xlrd import open_workbook
from xlutils.copy import copy
import shutil
import math
import os

from generateLabels import *
from generateCoC import *
from generatePackingList import *

def checkTable(df):
    df2 = df.dropna(how='all')
    # print(df)
    # print('-----------------------------------')
    # print(df2)
    if(df2.isnull().values.any()):
        return False, None
    return True, df2

def generateDocs(input, plist, coc, dest, label):
    ############################
    # Check sheet 1 for errors #
    ############################
    
    # Check for top parameters
    info = pd.read_excel(input, nrows=3, usecols = 'A, B', header=None)
    po_num = info.loc[0][1]
    if(pd.isna(po_num)):
        label['text'] = 'Missing "PO#" in sheet "Part #"'
        return 404
    batch = info.loc[1][1]
    if(pd.isna(batch)):
        label['text'] = 'Missing "Batch" in sheet "Part #"'
        return 404
    thedate = info.loc[2][1]
    if(pd.isna(thedate)):
        label['text'] = 'Missing "Date" in sheet "Part #"'
        return 404

    packs = pd.read_excel(input, skiprows = 6)

    # Check packs for missing parameters
    packlist = []
    i = 0
    count = 0
    while i < packs.shape[1]:
        count += 1
        print(count)
        pack = pd.read_excel(input, skiprows = 6, usecols = [i, i+1])
        (success, pack) = checkTable(pack)
        if(not success):
            label['text'] = 'Missing values in sheet "Part #"'
            return 404
        pack.columns = ['Code', 'S/N']
        packlist.append(pack)
        i += 3

    # All parts
    parts = pd.concat(packlist, ignore_index=True)

    ############################
    # Check sheet 2 for errors #
    ############################
    # Check top parameters
    properties = pd.read_excel(input, sheet_name = 'Property', nrows=3, usecols = 'A, B', header=None)
    ppty = properties.loc[0][1]
    if(pd.isna(ppty)):
        label['text'] = 'Missing "Property" in sheet "Property"'
        return 404
    unit = properties.loc[1][1]
    if(pd.isna(unit)):
        label['text'] = 'Missing "Unit" in sheet "Property"'
        return 404

    # Check table for missing values
    part_ppty = pd.read_excel(input, sheet_name = 'Property', skiprows=4)
    if(not checkTable(part_ppty)[0]):        
        label['text'] = 'Missing values in sheet "Property"'
        return 404

    ############################
    # Check sheet 3 for errors #
    ############################
    # Check for missing values
    summary = pd.read_excel(input, sheet_name = 'Summary', skiprows=2, 
                            usecols = [1,2,3,4])
    (summ_success, summary) = checkTable(summary)
    if(not summ_success):
        label['text'] = 'Missing values in sheet "Summary"'
        return 404
    
    ############################
    ### Check serial numbers ###
    ############################

    for pack in packlist:
        for i in range(len(pack)):
            code = str(parts.iloc[i].at['Code'])
            row, col = np.where(summary == code)
            if len(row) == 0:
                label['text'] = 'Missing serial number for ' + code
                return 404
            

    ############################
    #### Make the Documents ####
    ############################
    docs = []

    i = 0
    counter = 1
    for pack in packlist:
        filepath, doc = makeLabels(pack, po_num, batch, thedate, counter, dest, summary)
        docs.append((filepath, doc))
        counter += 1

    cocs = makeCoC(parts, batch, thedate, part_ppty, ppty, unit, 
                            po_num, coc, dest, summary)
    docs.extend(cocs)

    filepath, doc, updates, totals = makePList(parts, part_ppty, summary, ppty, 
                            batch, po_num, thedate, counter-1, plist, dest, 
                            label)
    docs.append((filepath, doc))
    
    for (filepath, doc) in docs:
        doc.save(filepath)

    app = xw.App(visible=False)

    # Copy template sheet into batch folder
    savepath = os.path.split(filepath)[0]
    # print(savepath)
    batchsheet = savepath + '\\' + str(batch) + '.xlsx'  

    if os.path.isfile(batchsheet):
        filename, extension = os.path.splitext(batchsheet)
        counter = 1
        while os.path.isfile(batchsheet):
            batchsheet = filename + " (" + str(counter) + ")" + extension
            counter += 1
            
    shutil.copy(input, batchsheet)

    bb = xw.Book(batchsheet)
    bs = bb.sheets['Summary']
    
    bb.sheets["Part #"].delete()
    bb.sheets["Property"].delete()

    # # Fill in the summary columns
    bs.range('F4').options(index=False, header=False).value = updates
    bb.save()
    bb.close()

    # Clear existing sheet for new template
    # wb = xw.Book(input)

    # Clear batch and date
    # wb.sheets['Part #'].range('B2').clear_contents()
    # wb.sheets['Part #'].range('B3').clear_contents()
    # wb.sheets['Part #'][7:,:].clear_contents()

    # wb.sheets['Property'][5:,:].clear_contents()

    # wb.sheets['Summary'][3:,4:].clear_contents()

    # wb.sheets['Summary'].range('E4').options(index=False, header=False).value = totals

    # wb.save()
    # wb.close()

    app.quit()
    
    return dest + '/' + batch