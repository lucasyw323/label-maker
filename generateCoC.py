from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
from docx import Document
import pandas as pd
import numpy as np
import os

#######################################
##### This section is for the CoC #####
#######################################

def makeCoC(parts, batch, date, part_ppty, ppty, unit, po_num, template, dest, summary):
    # Makes a dataframe where a code is matched to each of its S/N's
    unique_parts = parts.groupby('Code')['S/N'].apply(list).reset_index(name = 'S/N')

    numcodes = unique_parts.shape[0]

    docs = []

    for i in range(numcodes):
        # Generate description with only one code
        code_desc = ""
        code = unique_parts.loc[i].at['Code']
        sns = unique_parts.loc[i].at['S/N']
        start = -1000
        last = -1000
        counter = 0
        for sn in sorted(sns):
            if sn == last+1:
                if(counter == len(sns)-1):
                    code_desc += (' to ' + str(code) + '-' + str(last))
                pass
            else:
                if(start != last):      # Last code was consecutive
                    code_desc += (' to ' + str(code) + '-' + str(last))
                    code_desc += (', ' + str(code) + '-' + str(sn))
                else:
                    code_desc += (', ' + str(code) + '-' + str(sn))
                start = sn
            last = sn
            counter += 1
        
        code = unique_parts.loc[i].at['Code']
        sns = unique_parts.loc[i].at['S/N']
        context = dict()
        context['desc'] = code_desc[2:]
        context['qnty'] = len(sns)
        context['code'] = code
        context['PO'] = po_num
        sumRow, sumCol = np.where(summary == code)
        context['PN'] = summary[sumRow[0]][sumCol[0] - 1]

        # if code == 'M38':
        #     context['PN'] = '1110-038-2'
        # elif code == 'M39':
        #     context['PN'] = '1110-039-2'
        # elif code == 'M26':
        #     context['PN'] = '1120-026-2'
        # elif code == 'M30':
        #     context['PN'] = '1130-030-2'
        # elif code == 'M17':
        #     context['PN'] = '1141-017-2'
        # elif code == 'M18':
        #     context['PN'] = '1141-018-2'
        # elif code == 'M40':
        #     context['PN'] = '1142-040-2'
        # elif code == 'M37':
        #     context['PN'] = '1142-037-2'
        # elif code == 'M42':
        #     context['PN'] = '1142-042-2'
        # elif code == 'M43':
        #     context['PN'] = '1142-043-2'
        context['Batch'] = batch
        context['Date'] = str(date)[:10]
        context['Property'] = ppty
        context['Unit'] = unit

        context['QA_data'] = []

        count = 0
        for i in range(part_ppty.shape[0]):
            if(str(part_ppty.iloc[i][0]) != code):
                continue
            serial = part_ppty.iloc[i][1]
            value = part_ppty.iloc[i][2]
            context['QA_data'].append(((code + '-' + str(serial)), value))
            count += 1

        context['QAqnty'] = count

        CoC = DocxTemplate(template)
        CoC.render(context)

        savepath = dest + '\\' +  context['Batch']
        filepath = savepath + '\\CoC' + '_'+str(batch) +'_'+ str(code) +'.docx'

        if not os.path.exists(savepath):
            os.mkdir(savepath)

        if os.path.isfile(filepath):
            filename, extension = os.path.splitext(filepath)
            counter = 1
            while os.path.isfile(filepath):
                filepath = filename + " (" + str(counter) + ")" + extension
                counter += 1
            docs.append((filepath, CoC))
        else:
            docs.append((filepath, CoC))
    return docs