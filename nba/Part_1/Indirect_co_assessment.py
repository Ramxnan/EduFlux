from openpyxl.styles import Font, Alignment                                           #import font and alignment from openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo          #import table from openpyxl
from openpyxl.styles import Protection 
from openpyxl.utils import get_column_letter                                  #import get_column_letter from openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill
from openpyxl.styles import Border, Side
from openpyxl.formatting.rule import FormulaRule


def indirect_co_assessment(data,aw):
    #merge cells depending on number of POs
    aw.merge_cells(start_row=data["Number_of_COs"]+5, start_column=4, end_row=data["Number_of_COs"]+5, end_column=5)
    aw[f'D{data["Number_of_COs"]+5}']="Indirect CO Assessment"
    aw[f'D{data["Number_of_COs"]+5}'].font = Font(bold=True)
    aw[f'D{data["Number_of_COs"]+5}'].fill = PatternFill(start_color="ffe74e", end_color="ffe74e", fill_type="solid")
    for row in aw.iter_rows(min_row=data["Number_of_COs"]+5, max_row=data["Number_of_COs"]+5, min_col=4, max_col=5):
        for cell in row:
            cell.border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))

    aw[f"D{data['Number_of_COs']+6}"]="COs"
    aw[f"D{data['Number_of_COs']+6}"].font = Font(bold=True)
    aw[f"D{data['Number_of_COs']+6}"].fill = PatternFill(start_color='f79646', end_color='f79646', fill_type='solid')
    aw[f"D{data['Number_of_COs']+6}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))
    

    for i in range(1,data["Number_of_COs"]+1):
        aw[f"D{i+data['Number_of_COs']+6}"]=f"CO{i}"
        aw[f"D{i+data['Number_of_COs']+6}"].font = Font(bold=True)
        #border
        aw[f"D{i+data['Number_of_COs']+6}"].border = Border(top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'),
                                    left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'))
        aw[f"E{i+data['Number_of_COs']+6}"].border = Border(top=Side(border_style='thin', color='000000'),
                                    bottom=Side(border_style='thin', color='000000'),
                                    left=Side(border_style='thin', color='000000'),
                                    right=Side(border_style='thin', color='000000'))
        if i%2==0:
            aw[f"D{i+data['Number_of_COs']+6}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
            aw[f"E{i+data['Number_of_COs']+6}"].fill = PatternFill(start_color='fde9d9', end_color='fde9d9', fill_type='solid')
        else:
            aw[f"D{i+data['Number_of_COs']+6}"].fill = PatternFill(start_color='fcd5b4' , end_color='fcd5b4', fill_type='solid')
            aw[f"E{i+data['Number_of_COs']+6}"].fill = PatternFill(start_color='fcd5b4' , end_color='fcd5b4', fill_type='solid')

    aw[f"E{data['Number_of_COs']+6}"]="Indirect %"
    aw[f"E{data['Number_of_COs']+6}"].font = Font(bold=True)
    aw[f"E{data['Number_of_COs']+6}"].fill = PatternFill(start_color='f79646', end_color='f79646', fill_type='solid')
    aw[f"E{data['Number_of_COs']+6}"].border = Border(top=Side(border_style='thin', color='000000'),
                                bottom=Side(border_style='thin', color='000000'),
                                left=Side(border_style='thin', color='000000'),
                                right=Side(border_style='thin', color='000000'))


    # Create a data validation object for numbers between 0 and 100
    number_validation = DataValidation(type="decimal", operator="between", formula1=0, formula2=100)
    number_validation.error = 'You must enter a number between 0 and 100'
    number_validation.errorTitle = 'Invalid Entry'
    number_validation.showErrorMessage = True

    # Apply this validation to the specified cells
    for row in range(2+data["Number_of_COs"]+5, 2+data["Number_of_COs"]+5 + data["Number_of_COs"]):
        cell_coordinate = f'E{row}'
        aw[cell_coordinate].protection = Protection(locked=False)
        aw.add_data_validation(number_validation)
        number_validation.add(aw[cell_coordinate])

    pink_fill = PatternFill(start_color="D8A5B5", end_color="D8A5B5", fill_type="solid")
    for row in range(2+data["Number_of_COs"]+5, 2+data["Number_of_COs"]+5+data["Number_of_COs"]):
        cell_coordinate = f'E{row}'
        aw.conditional_formatting.add(
            cell_coordinate,
            FormulaRule(formula=[f'ISBLANK({cell_coordinate})'], stopIfTrue=False, fill=pink_fill)
        )

    


    return aw