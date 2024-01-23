from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries                                  #import get_column_letter from openpyxl
import time
from openpyxl import Workbook                                                         #import workbook from openpyxl
import pandas as pd
import os
import numpy as np

def driver_part3(input_dir_path, output_dir_path):
    wbfinal = Workbook()
    wbfinal.remove(wbfinal.active)
    wbfinal.create_sheet("Printouts",0)
    wbfinal.create_sheet("PO_calculation",1)
    ws=wbfinal["Printouts"]
    for file in os.listdir(input_dir_path):
        wb=load_workbook(input_dir_path+"\\"+file, data_only=True) 
        print("=====================================")
        print(wb.sheetnames)
        print(file)
        print("=====================================")
        ws1=wb["Printout"]

        data={}
        inputTable=ws1.tables['Inputinfo']
        # Iterate through the rows in the input table
        table_range=inputTable.ref
        min_col, min_row, max_col, max_row = range_boundaries(table_range)
        for row in range(min_row, max_row+1):
            data[ws1.cell(row=row, column=1).value]=ws1.cell(row=row, column=2).value

        min_row=4
        min_col=4
        max_row=4+data["Number_of_COs :"]-1
        max_col=17
        rowdata=[]
        for row in range(min_row, max_row+1):
                rowdata.append([])
                for col in range(min_col, max_col+1):
                    rowdata[-1].append(ws1.cell(row=row, column=col).value)
        Printouttable = pd.DataFrame(rowdata[1:], columns=rowdata[0])
        #print(Printouttable)

        #prin the table in the final workbook
        #ws=wbfinal["Printouts"]
        ws.append([file])
        for r in dataframe_to_rows(Printouttable, index=False, header=True):
            ws.append(r)
        ws.append([])

    ws=wbfinal["PO_calculation"]
    final_po_table=pd.DataFrame()
    for file in os.listdir(input_dir_path):
        wb=load_workbook(input_dir_path+"\\"+file, data_only=True) 
        ws1=wb["Course Level Attainment"]
        #WeightedPO
        potable=ws1.tables['WeightedPO']
        table_range=potable.ref
        min_col, min_row, max_col, max_row = range_boundaries(table_range)
        rowdata=[]
        for row in range(min_row, max_row+1):
                rowdata.append([])
                for col in range(min_col, max_col+1):
                    rowdata[-1].append(ws1.cell(row=row, column=col).value)
        potable = pd.DataFrame(rowdata[-1:], columns=rowdata[0])
        potable=potable.replace(0, np.nan)
        #add potable to final_po using pandas concat
        final_po_table=pd.concat([final_po_table,potable], axis=0)

    #print(final_po_table)
    lastrow=[]
    # Calculate average, excluding NaN values
    for col in final_po_table.columns:
        #print(col)
        if col == "COs\POs":
            lastrow.append("Average")
        else:
            # Convert the column to numeric, non-numeric values become NaN
            numeric_col = pd.to_numeric(final_po_table[col], errors='coerce')
            # Calculate the mean of the column, skipping NaN values
            mean_value = numeric_col.mean()
            lastrow.append(mean_value)

    # Append the last row with the averages to your DataFrame
    final_po_table.loc['Average'] = lastrow

    #print(final_po_table)

    # Add the DataFrame to the Excel sheet
    for r in dataframe_to_rows(final_po_table, index=False, header=False):
        ws.append(r)

    wbfinal.save(os.path.join(output_dir_path, "final.xlsx"))
    return os.path.join("final.xlsx")