from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries                                  #import get_column_letter from openpyxl
import time
from openpyxl import Workbook                                                         #import workbook from openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, NamedStyle, colors, Color, Fill, GradientFill, Font, Border, Side, Alignment, Protection
import pandas as pd
import os
import numpy as np
from printout import printout_template


def driver_part2(input_dir_path, output_dir_path):
    wbwrite = Workbook()
    wbwrite.remove(wbwrite.active)
    wbwrite.create_sheet("Printouts",0)
    wbwrite.create_sheet("PO_calculations",1)
    
    wswrite_printouts=wbwrite["Printouts"]
    startrow=2
    for file in os.listdir(input_dir_path):
        if file.endswith(".xlsx") and file != "final.xlsx":
            wbread=load_workbook(input_dir_path+"\\"+file, data_only=True) 
            print("=====================================")
            print(wbread.sheetnames)
            print(file)
            print("=====================================")
            wsread_input_detials=wbread["Input_Details"]
            Number_of_COs=wsread_input_detials["B8"].value

            wsread_printout=wbread["Printout"]
            min_row=1
            min_col=1
            max_row=3+Number_of_COs
            max_col=14
            rowdata=[]
            for row in range(min_row, max_row+1):
                    rowdata.append([])
                    for col in range(min_col, max_col+1):
                        rowdata[-1].append(wsread_printout.cell(row=row, column=col).value)
            Printouttable = pd.DataFrame(rowdata[1:], columns=rowdata[0])
            #print(Printouttable)

            #prin the table in the final workbook
            wswrite_printouts.append([file])
            for r in dataframe_to_rows(Printouttable, index=False, header=True):
                wswrite_printouts.append(r)
            wswrite_printouts.append([])

            wswrite_printouts.merge_cells(f"A{startrow-1}:N{startrow-1}")
            wswrite_printouts[f"A{startrow-1}"].fill = PatternFill(start_color='D39554', end_color='D39554', fill_type='solid')
            wswrite_printouts[f"A{startrow-1}"].font = Font(bold=True)
            wswrite_printouts[f"A{startrow-1}"].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            #style the table
            printout_template(wswrite_printouts, startrow)
          
            min_row=startrow+3
            min_col=1
            max_row=startrow+3+Number_of_COs-1
            max_col=14
            for row in range(min_row, max_row+1):
                for col in range(min_col, max_col+1):
                    wswrite_printouts.cell(row=row, column=col).border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
                    wswrite_printouts.cell(row=row, column=col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    wswrite_printouts.cell(row=row, column=col).font = Font(bold=True)

            wswrite_printouts.merge_cells(f"A{startrow+3}:A{startrow+3+Number_of_COs-1}")
            wswrite_printouts[f"A{startrow+3}"].font = Font(bold=True)
            wswrite_printouts[f"A{startrow+3}"].alignment = Alignment(horizontal='center', vertical='center', textRotation=90, wrap_text=True)
            wswrite_printouts[f"A{startrow+3}"].fill = PatternFill(start_color='1ed760', end_color='1ed760', fill_type='solid')
            for numco in range(Number_of_COs):
                if numco%2==0:
                    wswrite_printouts[f"B{startrow+3+numco}"].fill = PatternFill(start_color='ffff00', end_color='ffff00', fill_type='solid')
                
                wswrite_printouts[f"D{startrow+3+numco}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
                wswrite_printouts[f"D{startrow+3+numco}"].font = Font(bold=True, color="fe3400")

                wswrite_printouts[f"F{startrow+3+numco}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
                wswrite_printouts[f"F{startrow+3+numco}"].font = Font(bold=True, color="fe3400")

                wswrite_printouts[f"H{startrow+3+numco}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
                wswrite_printouts[f"H{startrow+3+numco}"].font = Font(bold=True, color="fe3400")

                wswrite_printouts[f"J{startrow+3+numco}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
                wswrite_printouts[f"J{startrow+3+numco}"].font = Font(bold=True, color="fe3400")

                wswrite_printouts[f"L{startrow+3+numco}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
                wswrite_printouts[f"L{startrow+3+numco}"].font = Font(bold=True, color="fe3400")

                wswrite_printouts[f"M{startrow+3+numco}"].fill = PatternFill(start_color='ffff00', end_color='ffff00', fill_type='solid')

            startrow=startrow+Number_of_COs+5



    # #PO calculation
    wswrite_POCalculation=wbwrite["PO_calculations"]
    wswrite_POCalculation["A1"]="Course Code"
    wswrite_POCalculation["A1"].font = Font(bold=True, color="ffffff")
    wswrite_POCalculation["A1"].fill = PatternFill(start_color='6CB266', end_color='6CB266', fill_type='solid')
    wswrite_POCalculation["A1"].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    wswrite_POCalculation["A1"].border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
    wswrite_POCalculation.column_dimensions['A'].width = 12

    for po in range(1, 13):
        wswrite_POCalculation.cell(row=1, column=po+1).value=f"PO{po}"
        wswrite_POCalculation.cell(row=1, column=po+1).font = Font(bold=True, color="ffffff")
        wswrite_POCalculation.cell(row=1, column=po+1).fill = PatternFill(start_color='6CB266', end_color='6CB266', fill_type='solid')
        wswrite_POCalculation.cell(row=1, column=po+1).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        wswrite_POCalculation.cell(row=1, column=po+1).border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))

    for pso in range(1, 6):
        wswrite_POCalculation.cell(row=1, column=pso+13).value=f"PSO{pso}"
        wswrite_POCalculation.cell(row=1, column=pso+13).font = Font(bold=True, color="ffffff")
        wswrite_POCalculation.cell(row=1, column=pso+13).fill = PatternFill(start_color='6CB266', end_color='6CB266', fill_type='solid')
        wswrite_POCalculation.cell(row=1, column=pso+13).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        wswrite_POCalculation.cell(row=1, column=pso+13).border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))

    final_po_table=pd.DataFrame()
    for file in os.listdir(input_dir_path):
        if file.endswith(".xlsx") and file != "final.xlsx":
            wbread=load_workbook(input_dir_path+"\\"+file, data_only=True)
            wsread_input_detials=wbread["Input_Details"]
            Number_of_COs=wsread_input_detials["B8"].value
            wsread_Course_Level_Attainment=wbread["Course_level_Attainment"]
            row=wsread_Course_Level_Attainment.max_row
            min_col=2
            max_col=19
            rowdata=[]
            for col in range(min_col, max_col+1):
                    rowdata.append(wsread_Course_Level_Attainment.cell(row=row, column=col).value)
            #print(rowdata)
            rowdata_df=pd.DataFrame(rowdata).T
            final_po_table = pd.concat([final_po_table,rowdata_df], axis=0)
    final_po_table=final_po_table.replace(0, np.nan)

    #print(final_po_table)
    lastrow=[]
    # Calculate average, excluding NaN values
    for col in final_po_table.columns:
        if col == 0:
            lastrow.append("Average")
        else:
            # Convert the column to numeric, non-numeric values become NaN
            numeric_col = pd.to_numeric(final_po_table[col], errors='coerce')
            # Calculate the mean of the column, skipping NaN values
            mean_value = numeric_col.mean()
            lastrow.append(mean_value)

    # Append the last row with the averages to your DataFrame
    final_po_table.loc['Average'] = lastrow
    final_po_table.reset_index(drop=True, inplace=True)

    #write the final_po_table to the excel sheet from A2 to the end
    row=2
    for _ in dataframe_to_rows(final_po_table, index=False, header=False):
        for c in range(1, 19):
            wswrite_POCalculation.cell(row=row, column=c).value=final_po_table.iloc[row-2, c-1]
            wswrite_POCalculation.cell(row=row, column=c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            wswrite_POCalculation.cell(row=row, column=c).border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            if row%2!=0:
                wswrite_POCalculation.cell(row=row, column=c).fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
            
        row=row+1
    row-=1
    for c in range(1, 19):
        wswrite_POCalculation.cell(row=row, column=c).fill = PatternFill(start_color='d99cac', end_color='d99cac', fill_type='solid')
        wswrite_POCalculation.cell(row=row, column=c).font = Font(bold=True)



    # print(final_po_table)







                














    # ws=wbwrite["PO_calculation"]
    # final_po_table=pd.DataFrame()
    # for file in os.listdir(input_dir_path):
    #     wb=load_workbook(input_dir_path+"\\"+file, data_only=True) 
    #     ws1=wb["Course Level Attainment"]
    #     #WeightedPO
    #     potable=ws1.tables['WeightedPO']
    #     table_range=potable.ref
    #     min_col, min_row, max_col, max_row = range_boundaries(table_range)
    #     rowdata=[]
    #     for row in range(min_row, max_row+1):
    #             rowdata.append([])
    #             for col in range(min_col, max_col+1):
    #                 rowdata[-1].append(ws1.cell(row=row, column=col).value)
    #     potable = pd.DataFrame(rowdata[-1:], columns=rowdata[0])
    #     potable=potable.replace(0, np.nan)
    #     #add potable to final_po using pandas concat
    #     final_po_table=pd.concat([final_po_table,potable], axis=0)

    # #print(final_po_table)
    # lastrow=[]
    # # Calculate average, excluding NaN values
    # for col in final_po_table.columns:
    #     #print(col)
    #     if col == "COs\POs":
    #         lastrow.append("Average")
    #     else:
    #         # Convert the column to numeric, non-numeric values become NaN
    #         numeric_col = pd.to_numeric(final_po_table[col], errors='coerce')
    #         # Calculate the mean of the column, skipping NaN values
    #         mean_value = numeric_col.mean()
    #         lastrow.append(mean_value)

    # # Append the last row with the averages to your DataFrame
    # final_po_table.loc['Average'] = lastrow

    # #print(final_po_table)

    # # Add the DataFrame to the Excel sheet
    # for r in dataframe_to_rows(final_po_table, index=False, header=False):
    #     ws.append(r)

    wbwrite.save(os.path.join(output_dir_path, "final.xlsx"))
    return os.path.join("final.xlsx")

if __name__ == "__main__":
    file1 = "C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_2\\2019_19CSE345_Computer system and architecture_A_Even.xlsx"
    file2="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_2\\2019_19MEE444_PCE_A_Even.xlsx"
    input_dir_path="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_2"
    output_dir_path="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_2"
    driver_part2(input_dir_path, output_dir_path)