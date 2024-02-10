import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.styles import Protection

def write_course_attainment(data,Component_Details,aw):
    start_row_ca = 1
    start_col_ca = 5

    aw.merge_cells(start_row=start_row_ca, start_column=start_col_ca, end_row=start_row_ca+3, end_column=start_col_ca)
    aw[f'{get_column_letter(start_col_ca)}{start_row_ca}'] = 'Course Outcome'
    aw[f'{get_column_letter(start_col_ca)}{start_row_ca}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca, start_column=start_col_ca+1, end_row=start_row_ca, end_column=start_col_ca+2)
    aw[f'{get_column_letter(start_col_ca+1)}{start_row_ca}'] = 'Mapping with Program'
    aw[f'{get_column_letter(start_col_ca+1)}{start_row_ca}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+1, start_column=start_col_ca+1, end_row=start_row_ca+3, end_column=start_col_ca+1)
    aw[f'{get_column_letter(start_col_ca+1)}{start_row_ca+1}'] = 'POs & PSOs'
    aw[f'{get_column_letter(start_col_ca+1)}{start_row_ca+1}'].font = Font(bold=True)

    aw[f'{get_column_letter(start_col_ca+2)}{start_row_ca+1}'] = 'Level of Mapping'
    aw[f'{get_column_letter(start_col_ca+2)}{start_row_ca+1}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+2, start_column=start_col_ca+2, end_row=start_row_ca+3, end_column=start_col_ca+2)
    aw[f'{get_column_letter(start_col_ca+2)}{start_row_ca+2}'] = 'Affinity'
    aw[f'{get_column_letter(start_col_ca+2)}{start_row_ca+2}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca, start_column=start_col_ca+3, end_row=start_row_ca, end_column=start_col_ca+10)
    aw[f'{get_column_letter(start_col_ca+3)}{start_row_ca}'] = 'Attainment % in'
    aw[f'{get_column_letter(start_col_ca+3)}{start_row_ca}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+1, start_column=start_col_ca+3, end_row=start_row_ca+1, end_column=start_col_ca+7)
    aw[f'{get_column_letter(start_col_ca+3)}{start_row_ca+1}'] = 'Direct'
    aw[f'{get_column_letter(start_col_ca+3)}{start_row_ca+1}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+2, start_column=start_col_ca+3, end_row=start_row_ca+2, end_column=start_col_ca+4)
    aw[f'{get_column_letter(start_col_ca+3)}{start_row_ca+2}'] = 'University(SEE)'
    aw[f'{get_column_letter(start_col_ca+3)}{start_row_ca+2}'].font = Font(bold=True)

    aw[f'{get_column_letter(start_col_ca+3)}{start_row_ca+3}'] = 'Attainment'
    aw[f'{get_column_letter(start_col_ca+3)}{start_row_ca+3}'].font = Font(bold=True)

    aw[f'{get_column_letter(start_col_ca+4)}{start_row_ca+3}'] = 'Level Of Attainment (0-40 --> 1, 40-60 ---> 2, 60-100---> 3)'
    aw[f'{get_column_letter(start_col_ca+4)}{start_row_ca+3}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+2, start_column=start_col_ca+5, end_row=start_row_ca+2, end_column=start_col_ca+6)
    aw[f'{get_column_letter(start_col_ca+5)}{start_row_ca+2}'] = 'Internal(CIE)'
    aw[f'{get_column_letter(start_col_ca+5)}{start_row_ca+2}'].font = Font(bold=True)

    aw[f'{get_column_letter(start_col_ca+5)}{start_row_ca+3}'] = 'Attainment'
    aw[f'{get_column_letter(start_col_ca+5)}{start_row_ca+3}'].font = Font(bold=True)

    aw[f'{get_column_letter(start_col_ca+6)}{start_row_ca+3}'] = 'Level Of Attainment (0-40 --> 1, 40-60 ---> 2, 60-100---> 3)'
    aw[f'{get_column_letter(start_col_ca+6)}{start_row_ca+3}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+2, start_column=start_col_ca+7, end_row=start_row_ca+3, end_column=start_col_ca+7)
    aw[f'{get_column_letter(start_col_ca+7)}{start_row_ca+2}'] = 'Weighted Level of Attainment (University + IA)'
    aw[f'{get_column_letter(start_col_ca+7)}{start_row_ca+2}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+1, start_column=start_col_ca+8, end_row=start_row_ca+1, end_column=start_col_ca+9)
    aw[f'{get_column_letter(start_col_ca+8)}{start_row_ca+1}'] = 'Indirect'
    aw[f'{get_column_letter(start_col_ca+8)}{start_row_ca+1}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+2, start_column=start_col_ca+8, end_row=start_row_ca+3, end_column=start_col_ca+8)
    aw[f'{get_column_letter(start_col_ca+8)}{start_row_ca+2}'] = 'Attainment'
    aw[f'{get_column_letter(start_col_ca+8)}{start_row_ca+2}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+2, start_column=start_col_ca+9, end_row=start_row_ca+3, end_column=start_col_ca+9)
    aw[f'{get_column_letter(start_col_ca+9)}{start_row_ca+2}'] = 'Level Of Attainment'
    aw[f'{get_column_letter(start_col_ca+9)}{start_row_ca+2}'].font = Font(bold=True)

    aw.merge_cells(start_row=start_row_ca+1, start_column=start_col_ca+10, end_row=start_row_ca+2, end_column=start_col_ca+10)
    aw[f'{get_column_letter(start_col_ca+10)}{start_row_ca+1}'] = 'Final Weighted CO Attainment (80% Direct + 20% Indirect)'
    aw[f'{get_column_letter(start_col_ca+10)}{start_row_ca+1}'].font = Font(bold=True)

    aw[f'{get_column_letter(start_col_ca+10)}{start_row_ca+3}'] = 'Level of Attainment'
    aw[f'{get_column_letter(start_col_ca+10)}{start_row_ca+3}'].font = Font(bold=True)

    aw.column_dimensions[get_column_letter(start_col_ca)].width = 17.22
    aw.column_dimensions[get_column_letter(start_col_ca+1)].width = 9.33
    aw.column_dimensions[get_column_letter(start_col_ca+2)].width = 15.56
    aw.column_dimensions[get_column_letter(start_col_ca+3)].width = 12
    aw.column_dimensions[get_column_letter(start_col_ca+4)].width = 14.11
    aw.column_dimensions[get_column_letter(start_col_ca+5)].width = 12
    aw.column_dimensions[get_column_letter(start_col_ca+6)].width = 14.11
    aw.column_dimensions[get_column_letter(start_col_ca+7)].width = 20.67
    aw.column_dimensions[get_column_letter(start_col_ca+8)].width = 12
    aw.column_dimensions[get_column_letter(start_col_ca+9)].width = 18.11
    aw.column_dimensions[get_column_letter(start_col_ca+10)].width = 22.78
    
    for row in aw.iter_rows(min_row=start_row_ca, max_row=start_row_ca+3, min_col=start_col_ca, max_col=aw.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.fill = PatternFill(start_color='b8cce4', end_color='b8cce4', fill_type='solid')
            cell.border = Border(left=Side(border_style='thin', color='000000'),
                                 right=Side(border_style='thin', color='000000'),
                                 top=Side(border_style='thin', color='000000'),
                                 bottom=Side(border_style='thin', color='000000'))
            
    aw[f'{get_column_letter(start_col_ca+2)}{start_row_ca+2}'].fill = PatternFill(start_color='c4d79b', end_color='c4d79b', fill_type='solid')
    aw[f'{get_column_letter(start_col_ca+10)}{start_row_ca+3}'].fill = PatternFill(start_color='c4d79b', end_color='c4d79b', fill_type='solid')

    start=4
    interval=16
    rowindex=1
    for i in range(1, (data["Number_of_COs"]+1)):
        start+=1
        aw.merge_cells(start_row=start, start_column=1, end_row=start+interval, end_column=1)
        aw.cell(row=start, column=1).value = "CO"+str(i)
        aw.cell(row=start, column=1).font = Font(bold=True)
        aw.cell(row=start, column=1).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        if i%2==0:
            aw.cell(row=start, column=1).fill = PatternFill(start_color='c4d79b', end_color='c4d79b', fill_type='solid')
        else:
            aw.cell(row=start, column=1).fill = PatternFill(start_color='b8cce4', end_color='b8cce4', fill_type='solid')


        index=1
        for j in range(start, start+interval+1):
            #print out COPO mapping
            aw.cell(row=j, column=2).value =f'={data["Section"]}_Input_Details!{get_column_letter(index+4)}2'
            aw.cell(row=j, column=2).alignment = Alignment(horizontal='center', vertical='center')
            aw.cell(row=j, column=2).border = Border(left=Side(border_style='thin', color='000000'),
                                 right=Side(border_style='thin', color='000000'),
                                 top=Side(border_style='thin', color='000000'),
                                 bottom=Side(border_style='thin', color='000000'))
            
            aw.cell(row=j, column=3).value = f'={data["Section"]}_Input_Details!{get_column_letter(index+4)}{i+2}'
            aw.cell(row=j, column=3).alignment = Alignment(horizontal='center', vertical='center')
            aw.cell(row=j, column=3).border = Border(left=Side(border_style='thin', color='000000'),
                                 right=Side(border_style='thin', color='000000'),
                                 top=Side(border_style='thin', color='000000'),
                                 bottom=Side(border_style='thin', color='000000'))
            
            if index%2==0:
                aw.cell(row=j, column=2).fill = PatternFill(start_color='c4d79b', end_color='c4d79b', fill_type='solid')
                aw.cell(row=j, column=3).fill = PatternFill(start_color='c4d79b', end_color='c4d79b', fill_type='solid')
            else:
                aw.cell(row=j, column=2).fill = PatternFill(start_color='b8cce4', end_color='b8cce4', fill_type='solid')
                aw.cell(row=j, column=3).fill = PatternFill(start_color='b8cce4', end_color='b8cce4', fill_type='solid')
            index+=1

        for k in range(4,12):
            aw.merge_cells(start_row=start, start_column=k, end_row=start+interval, end_column=k)
            #aw.cell(row=start, column=k).value = final_table.iloc[i-1, k-4]
            aw.cell(row=start, column=k).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if k%2==0:
                aw.cell(row=start, column=k).fill = PatternFill(start_color='b8cce4', end_color='b8cce4', fill_type='solid')
            else:
                aw.cell(row=start, column=k).fill = PatternFill(start_color='dce6f1', end_color='dce6f1', fill_type='solid')
    
        internal_components_num=0
        external_components_num=0
        for component_name in Component_Details.keys():
            if component_name[-1]=="I":
                internal_components_num+=1
            else:
                external_components_num+=1
        col=(data["Number_of_COs"]*internal_components_num) + (1*internal_components_num) + 2 + rowindex
        row=(6 + data["Number_of_Students"] + 5)
        aw.cell(row=start, column=4).value=f'=IFERROR({data["Section"]}_Internal_Components!{get_column_letter(col)}{row}, 0)'

        aw.cell(row=start, column=5).value=f'=IF(AND({get_column_letter(4)}{start}>=0,{get_column_letter(4)}{start}<40),1,IF(AND({get_column_letter(4)}{start}>=40,{get_column_letter(4)}{start}<60),2,IF(AND({get_column_letter(4)}{start}>=60,{get_column_letter(4)}{start}<=100),3,"0")))'

        col=(data["Number_of_COs"]*external_components_num) + (1*external_components_num) + 2 + rowindex
        row=(6 + data["Number_of_Students"] + 5)
        aw.cell(row=start, column=6).value=f'=IFERROR({data["Section"]}_External_Components!{get_column_letter(col)}{row}, 0)'

        aw.cell(row=start, column=7).value=f'=IF(AND({get_column_letter(6)}{start}>=0,{get_column_letter(6)}{start}<40),1,IF(AND({get_column_letter(6)}{start}>=40,{get_column_letter(6)}{start}<60),2,IF(AND({get_column_letter(6)}{start}>=60,{get_column_letter(6)}{start}<=100),3,"0")))'

        column_5_cell = f'{get_column_letter(5)}{start}'
        column_7_cell = f'{get_column_letter(7)}{start}'
        calculation = f'{column_5_cell}*({data["Section"]}_Input_Details!B16/100)+{column_7_cell}*{data["Section"]}_Input_Details!B15/100'

        formula = f'={calculation}'
        aw.cell(row=start, column=8).value = formula

        aw.cell(row=start, column=9).value=f'=IF({data["Section"]}_Input_Details!E{2+data["Number_of_COs"]+4+rowindex}>0,{data["Section"]}_Input_Details!E{2+data["Number_of_COs"]+4+rowindex},"0")'

        aw.cell(row=start, column=10).value=f'=IF(AND({get_column_letter(9)}{start}>=0,{get_column_letter(9)}{start}<40),1,IF(AND({get_column_letter(9)}{start}>=40,{get_column_letter(9)}{start}<60),2,IF(AND({get_column_letter(9)}{start}>=60,{get_column_letter(9)}{start}<=100),3,"0")))'

        aw.cell(row=start, column=11).value=f'=({get_column_letter(8)}{start}*({data["Section"]}_Input_Details!B17/100))+({get_column_letter(10)}{start}*({data["Section"]}_Input_Details!B18/100))'
        
        rowindex+=1
        start=start+interval

    for row in aw.iter_rows(min_row=1, max_row=aw.max_row, min_col=1, max_col=aw.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(left=Side(border_style='thin', color='000000'),
                                 right=Side(border_style='thin', color='000000'),
                                 top=Side(border_style='thin', color='000000'),
                                 bottom=Side(border_style='thin', color='000000'))
            

    #================================================================================================================================================================
    current_row = aw.max_row+4
    current_col = 4
    aw.merge_cells(start_row=current_row, start_column=current_col, end_row=current_row, end_column=17+current_col)
    aw.cell(row=current_row, column=current_col).value = "Weighted PO/PSO Attainment Contribution"
    aw.cell(row=current_row, column=current_col).font = Font(bold=True)
    aw.cell(row=current_row, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row, column=current_col).fill = PatternFill(start_color='E6BA62', end_color='E6BA62', fill_type='solid')
    

    current_row+=1
    aw.cell(row=current_row, column=current_col).value = "COs\\POs"
    aw.cell(row=current_row, column=current_col).font = Font(bold=True, color="FFFFFF")
    aw.cell(row=current_row, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    aw.cell(row=current_row, column=current_col).fill = PatternFill(start_color='4bacc6', end_color='4bacc6', fill_type='solid')

    for po in range(1,12+1):
        aw[f"{get_column_letter(po+current_col)}{current_row}"]=f"PO{po}"
        aw[f"{get_column_letter(po+current_col)}{current_row}"].font = Font(bold=True, color="FFFFFF")
        aw[f"{get_column_letter(po+current_col)}{current_row}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"{get_column_letter(po+current_col)}{current_row}"].border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
        aw[f"{get_column_letter(po+current_col)}{current_row}"].fill = PatternFill(start_color='4bacc6', end_color='4bacc6', fill_type='solid')
        
    
    for pso in range(1,6):
        aw[f"{get_column_letter(12+current_col+pso)}{current_row}"]=f"PSO{pso}"
        aw[f"{get_column_letter(12+current_col+pso)}{current_row}"].font = Font(bold=True, color="FFFFFF")
        aw[f"{get_column_letter(12+current_col+pso)}{current_row}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"{get_column_letter(12+current_col+pso)}{current_row}"].border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
        aw[f"{get_column_letter(12+current_col+pso)}{current_row}"].fill = PatternFill(start_color='4bacc6', end_color='4bacc6', fill_type='solid')
        

    current_row+=1

    start=4
    interval=17

    current_col=4
    for co in range(1,data["Number_of_COs"]+1):
        aw[f"{get_column_letter(current_col)}{current_row}"]=f"CO{co}"
        aw[f"{get_column_letter(current_col)}{current_row}"].font = Font(bold=True, color="FFFFFF")
        aw[f"{get_column_letter(current_col)}{current_row}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"{get_column_letter(current_col)}{current_row}"].border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
        aw[f"{get_column_letter(current_col)}{current_row}"].fill = PatternFill(start_color='4bacc6', end_color='4bacc6', fill_type='solid')

        for po in range(1,12+1):
            aw[f"{get_column_letter(po+current_col)}{current_row}"]=f'=C{start+po}*K{start+1}'
            aw[f"{get_column_letter(po+current_col)}{current_row}"].alignment = Alignment(horizontal='center', vertical='center')
            aw[f"{get_column_letter(po+current_col)}{current_row}"].border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
            
            

        for pso in range(1,6):
            aw[f"{get_column_letter(12+current_col+pso)}{current_row}"]=f'=C{start+12+pso}*K{start+1}'
            aw[f"{get_column_letter(12+current_col+pso)}{current_row}"].alignment = Alignment(horizontal='center', vertical='center')
            aw[f"{get_column_letter(12+current_col+pso)}{current_row}"].border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))

        current_row+=1
        start+=interval

    current_col=1
    aw.cell(row=current_row, column=current_col).value = "Academic Year"
    aw.cell(row=current_row, column=current_col).font = Font(bold=True)
    aw.cell(row=current_row, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row, column=current_col).fill = PatternFill(start_color='E6BA62', end_color='E6BA62', fill_type='solid')
    aw.cell(row=current_row, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    aw.cell(row=current_row+1, column=current_col).value = data["Academic_year"]
    aw.cell(row=current_row+1, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row+1, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    aw.cell(row=current_row+1, column=current_col).fill = PatternFill(start_color='4bacc6', end_color='4bacc6', fill_type='solid')
    aw.cell(row=current_row+1, column=current_col).font = Font(bold=True, color="FFFFFF")

    current_col+=1
    aw.cell(row=current_row, column=current_col).value = "Semester"
    aw.cell(row=current_row, column=current_col).font = Font(bold=True)
    aw.cell(row=current_row, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row, column=current_col).fill = PatternFill(start_color='E6BA62', end_color='E6BA62', fill_type='solid')
    aw.cell(row=current_row, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    aw.cell(row=current_row+1, column=current_col).value = data["Semester"]
    aw.cell(row=current_row+1, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row+1, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    aw.cell(row=current_row+1, column=current_col).fill = PatternFill(start_color='4bacc6', end_color='4bacc6', fill_type='solid')
    aw.cell(row=current_row+1, column=current_col).font = Font(bold=True, color="FFFFFF")
    current_col+=1

    aw.cell(row=current_row, column=current_col).value = "Subject Name"
    aw.cell(row=current_row, column=current_col).font = Font(bold=True)
    aw.cell(row=current_row, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row, column=current_col).fill = PatternFill(start_color='E6BA62', end_color='E6BA62', fill_type='solid')
    aw.cell(row=current_row, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    aw.cell(row=current_row+1, column=current_col).value = data["Subject_Name"]
    aw.cell(row=current_row+1, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row+1, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    aw.cell(row=current_row+1, column=current_col).fill = PatternFill(start_color='4bacc6', end_color='4bacc6', fill_type='solid')
    aw.cell(row=current_row+1, column=current_col).font = Font(bold=True, color="FFFFFF")
    current_col+=1

    aw.cell(row=current_row, column=current_col).value = "Subject Code"
    aw.cell(row=current_row, column=current_col).font = Font(bold=True)
    aw.cell(row=current_row, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row, column=current_col).fill = PatternFill(start_color='E6BA62', end_color='E6BA62', fill_type='solid')
    aw.cell(row=current_row, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    
    current_col+=1





    aw.merge_cells(start_row=current_row, start_column=current_col, end_row=current_row, end_column=16+current_col)
    aw.cell(row=current_row, column=current_col).value = "Final Ratio"
    aw.cell(row=current_row, column=current_col).font = Font(bold=True)
    aw.cell(row=current_row, column=current_col).alignment = Alignment(horizontal='center', vertical='center')
    aw.cell(row=current_row, column=current_col).fill = PatternFill(start_color='E6BA62', end_color='E6BA62', fill_type='solid')
    aw.cell(row=current_row, column=current_col).border = Border(left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'),
                                    top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'))
    

    current_row+=1
    current_col=4
    aw[f"{get_column_letter(current_col)}{current_row}"]=f"{data['Subject_Code']}"
    aw[f"{get_column_letter(current_col)}{current_row}"].font = Font(bold=True, color="FFFFFF")
    aw[f"{get_column_letter(current_col)}{current_row}"].alignment = Alignment(horizontal='center', vertical='center')
    aw[f"{get_column_letter(current_col)}{current_row}"].border = Border(left=Side(border_style='thin', color='000000'),
                                 right=Side(border_style='thin', color='000000'),
                                 top=Side(border_style='thin', color='000000'),
                                 bottom=Side(border_style='thin', color='000000'))
    aw[f"{get_column_letter(current_col)}{current_row}"].fill = PatternFill(start_color='4bacc6', end_color='4bacc6', fill_type='solid')

    for po in range(1,12+1):
        main_formula = f'SUM({get_column_letter(po+current_col)}{current_row-1-data["Number_of_COs"]}:{get_column_letter(po+current_col)}{current_row-2})/(SUM({data["Section"]}_Input_Details!{get_column_letter(4+po)}3:{data["Section"]}_Input_Details!{get_column_letter(4+po)}{data["Number_of_COs"]+2}))'
        complete_formula = f'=IF(AND(SUM({get_column_letter(po+current_col)}{current_row-1-data["Number_of_COs"]}:{get_column_letter(po+current_col)}{current_row-2})>0, SUM({data["Section"]}_Input_Details!{get_column_letter(4+po)}3:{data["Section"]}_Input_Details!{get_column_letter(4+po)}{data["Number_of_COs"]+2})>0), {main_formula}, 0)'
        aw[f"{get_column_letter(po+current_col)}{current_row}"] = complete_formula
        aw[f"{get_column_letter(po+current_col)}{current_row}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"{get_column_letter(po+current_col)}{current_row}"].border = Border(left=Side(border_style='thin', color='000000'),
                                 right=Side(border_style='thin', color='000000'),
                                 top=Side(border_style='thin', color='000000'),
                                 bottom=Side(border_style='thin', color='000000'))

    for pso in range(1,6):
        main_formula = f'SUM({get_column_letter(12+current_col+pso)}{current_row-1-data["Number_of_COs"]}:{get_column_letter(12+current_col+pso)}{current_row-2})/(SUM({data["Section"]}_Input_Details!{get_column_letter(12+4+pso)}3:{data["Section"]}_Input_Details!{get_column_letter(12+4+pso)}{data["Number_of_COs"]+2}))'
        complete_formula = f'=IF(AND(SUM({get_column_letter(12+current_col+pso)}{current_row-1-data["Number_of_COs"]}:{get_column_letter(12+current_col+pso)}{current_row-2})>0, SUM({data["Section"]}_Input_Details!{get_column_letter(12+4+pso)}3:{data["Section"]}_Input_Details!{get_column_letter(12+4+pso)}{data["Number_of_COs"]+2})>0), {main_formula}, 0)'
        aw[f"{get_column_letter(12+current_col+pso)}{current_row}"] = complete_formula
        aw[f"{get_column_letter(12+current_col+pso)}{current_row}"].alignment = Alignment(horizontal='center', vertical='center')
        aw[f"{get_column_letter(12+current_col+pso)}{current_row}"].border = Border(left=Side(border_style='thin', color='000000'),
                                 right=Side(border_style='thin', color='000000'),
                                 top=Side(border_style='thin', color='000000'),
                                 bottom=Side(border_style='thin', color='000000'))

    current_row+=1


    for row in aw.iter_rows():
        for cell in row:
            cell.protection = Protection(locked=True)
    aw.protection.sheet = True

    

    return aw