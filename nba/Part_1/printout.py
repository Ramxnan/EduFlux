from openpyxl.styles import Font, Alignment                                           #import font and alignment from openpyxl
from openpyxl.styles.borders import Border, Side                                #import border from openpyxl
from openpyxl.styles import PatternFill                                           #import patternfill from openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

from .utils import adjust_width
from openpyxl.styles import Protection

import numpy as np

def printout_template(aw, data):
    aw.merge_cells("A1:A3")
    aw["A1"]="Course"
    aw["A1"].font = Font(bold=True)
    aw["A1"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=1, max_row=3, min_col=1, max_col=1):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
            

    aw.merge_cells("B1:B3")
    aw["B1"]="COs"
    aw["B1"].font = Font(bold=True)
    aw["B1"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=1, max_row=3, min_col=2, max_col=2):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
            
    aw.merge_cells("C1:D1")
    aw["C1"]="End Semester Examination"
    aw["C1"].font = Font(bold=True)
    aw["C1"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=1, max_row=1, min_col=3, max_col=4):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

    aw.merge_cells("C2:D2")
    aw["C2"]="(SEE)*"
    aw["C2"].font = Font(bold=True)
    aw["C2"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=2, max_row=2, min_col=3, max_col=4):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

    aw["C3"]="Attainment"
    aw["C3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["C3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    aw["D3"]="Level"
    aw["D3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["D3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))

    aw.merge_cells("E1:F1")
    aw["E1"]="Internal Examination"
    aw["E1"].font = Font(bold=True)
    aw["E1"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=1, max_row=1, min_col=5, max_col=6):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

    aw.merge_cells("E2:F2")
    aw["E2"]="(CIE)*"
    aw["E2"].font = Font(bold=True)
    aw["E2"].alignment = Alignment(horizontal='center', vertical='center')

    aw["E3"]="Attainment"
    aw["E3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["E3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    aw["F3"]="Level"
    aw["F3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["F3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    

    aw.merge_cells("G1:H1")
    aw["G1"]="Direct"
    aw["G1"].font = Font(bold=True)
    aw["G1"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=1, max_row=1, min_col=7, max_col=8):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
            
    aw.merge_cells("G2:H2")
# Formula to concatenate the header text with values from Input_Details!B15 and B16
    aw["G2"].value = f"=Input_Details!B15 & \" % of CIE + \" & Input_Details!B16 & \" % of SEE\""
    aw["G2"].font = Font(bold=True)
    aw["G2"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=2, max_row=2, min_col=7, max_col=8):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

    aw["G3"]="Attainment"
    aw["G3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["G3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    aw["H3"]="Level"
    aw["H3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["H3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))

    aw.merge_cells("I1:J2")
    aw["I1"]="Indirect"
    aw["I1"].font = Font(bold=True)
    aw["I1"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=1, max_row=2, min_col=9, max_col=10):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
            
    aw["I3"]="Attainment"
    aw["I3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["I3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    aw["J3"]="Level"
    aw["J3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["J3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))

    aw.merge_cells("K1:L1")
    aw["K1"]="Total Course Attainment"
    aw["K1"].font = Font(bold=True)
    aw["K1"].alignment = Alignment(horizontal='center', vertical='center')
    for row in aw.iter_rows(min_row=1, max_row=1, min_col=11, max_col=12):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
            
    aw.merge_cells("K2:L2")
# Formula to concatenate the header text with values from Input_Details!B17 and B18
    aw["K2"].value = f"=Input_Details!B17 & \" % of Direct + \" & Input_Details!B18 & \" % of Indirect\""
    aw["K2"].font = Font(bold=True)
    aw["K2"].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for row in aw.iter_rows(min_row=2, max_row=2, min_col=11, max_col=12):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

    aw["K3"]="Attainment"
    aw["K3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["K3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    aw["L3"]="Level"
    aw["L3"].alignment = Alignment(horizontal='center', vertical='center')
    aw["L3"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))

    aw["M1"]="Target"
    aw["M1"].font = Font(bold=True)
    aw["M1"].alignment = Alignment(horizontal='center', vertical='center')
    aw["M1"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
            
    aw["M2"]="(%)"
    aw["M2"].font = Font(bold=True)
    aw["M2"].alignment = Alignment(horizontal='center', vertical='center')
    aw["M2"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    
    aw["N1"]="Attainment"
    aw["N1"].font = Font(bold=True)
    aw["N1"].alignment = Alignment(horizontal='center', vertical='center')
    aw["N1"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    
    aw["N2"]="Yes/No"
    aw["N2"].font = Font(bold=True)
    aw["N2"].alignment = Alignment(horizontal='center', vertical='center')
    aw["N2"].border = Border(top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),
                            left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'))
    
    
    aw.column_dimensions['A'].width = 8.43
    aw.column_dimensions['B'].width = 8.43
    aw.column_dimensions['C'].width = 12
    aw.column_dimensions['D'].width = 12
    aw.column_dimensions['E'].width = 12
    aw.column_dimensions['F'].width = 12
    aw.column_dimensions['G'].width = 12
    aw.column_dimensions['H'].width = 12
    aw.column_dimensions['I'].width = 12
    aw.column_dimensions['J'].width = 8.43
    aw.column_dimensions['K'].width = 20
    aw.column_dimensions['L'].width = 8.43
    aw.column_dimensions['M'].width = 8.43
    aw.column_dimensions['N'].width = 10

    
    for row in aw.iter_cols(min_row=3, max_row=3, min_col=3, max_col=14):
        for cell in row:
            cell.fill = PatternFill(start_color='8db4e2', end_color='8db4e2', fill_type='solid')



def printout(aw, data):
    printout_template(aw,data)
    #merge D4 to number of COs
    aw.merge_cells(f"A4:A{3+data['Number_of_COs']}")
    #write the course name horizontally
    aw["A4"]=data["Subject_Name"]
    aw["A4"].font = Font(bold=True)
    aw["A4"].alignment = Alignment(horizontal='center', vertical='center', textRotation=90, wrap_text=True)
    aw["A4"].fill = PatternFill(start_color='1ed760', end_color='1ed760', fill_type='solid')

    start=4
    interval=17
    for i in range(data["Number_of_COs"]):
        aw[f"B{4+i}"]=f"CO{i+1}"
        aw[f"B{4+i}"].font = Font(bold=True)
        aw[f"B{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        if i%2==0:
            aw[f"B{4+i}"].fill = PatternFill(start_color='ffff00', end_color='ffff00', fill_type='solid')
        aw[f"B{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        
        aw[f"C{4+i}"]=f"=Course_Level_Attainment!D{5+(i*interval)}"
        aw[f"C{4+i}"].font = Font(bold=True)
        aw[f"C{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"C{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

        aw[f"D{4+i}"]=f"=Course_level_Attainment!E{5+(i*interval)}"
        aw[f"D{4+i}"].font = Font(bold=True, color="fe3400")
        aw[f"D{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"D{4+i}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
        aw[f"D{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        
        aw[f"E{4+i}"]=f"=Course_level_Attainment!F{5+(i*interval)}"
        aw[f"E{4+i}"].font = Font(bold=True)
        aw[f"E{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"E{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

        aw[f"F{4+i}"]=f"=Course_level_Attainment!G{5+(i*interval)}"
        aw[f"F{4+i}"].font = Font(bold=True, color="fe3400")
        aw[f"F{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"F{4+i}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
        aw[f"F{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

        aw[f"G{4+i}"]=f"=C{4+i}*(Input_Details!B16/100)+E{4+i}*(Input_Details!B15/100)"
        aw[f"G{4+i}"].font = Font(bold=True)
        aw[f"G{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"G{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        
        aw[f"H{4+i}"]=f"=Course_level_Attainment!H{5+(i*interval)}"
        aw[f"H{4+i}"].font = Font(bold=True, color="fe3400")
        aw[f"H{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"H{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        aw[f"H{4+i}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
        
        aw[f"I{4+i}"]=f"=Course_level_Attainment!I{5+(i*interval)}"
        aw[f"I{4+i}"].font = Font(bold=True)
        aw[f"I{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"I{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        
        aw[f"J{4+i}"]=f"=Course_level_Attainment!J{5+(i*interval)}"
        aw[f"J{4+i}"].font = Font(bold=True, color="fe3400")
        aw[f"J{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"J{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        aw[f"J{4+i}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
        
        aw[f"K{4+i}"]=f"=(G{4+i}*(Input_Details!B17/100))+(I{4+i}*(Input_Details!B18/100))"
        aw[f"K{4+i}"].font = Font(bold=True)
        aw[f"K{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"K{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        
        aw[f"L{4+i}"]=f"=Course_level_Attainment!K{5+(i*interval)}"
        aw[f"L{4+i}"].font = Font(bold=True, color="fe3400")
        aw[f"L{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"L{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        aw[f"L{4+i}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
        
        aw[f"M{4+i}"]=f"=Input_Details!B19"
        aw[f"M{4+i}"].font = Font(bold=True)
        aw[f"M{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"M{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
        aw[f"M{4+i}"].fill = PatternFill(start_color='ffff00', end_color='ffff00', fill_type='solid')
        
        
        aw[f"N{4+i}"]=f'=IF(K{4+i}>=M{4+i},"Yes","No")'
        aw[f"N{4+i}"].font = Font(bold=True)
        aw[f"N{4+i}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"N{4+i}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
    for row in aw.iter_rows():
        for cell in row:
            cell.protection = Protection(locked=True)
    aw.protection.sheet = True
    return aw