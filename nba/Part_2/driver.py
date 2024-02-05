from openpyxl import Workbook                                                         #import workbook from openpyxl
from .utils import adjust_width
from .Input_detail_table import input_detail
from .CO_PO_Table import CO_PO_Table
from .Indirect_co_assessment import indirect_co_assessment
from .qn_co_mm_btl import qn_co_mm_btl
from .studentmarks import studentmarks
from .cummulative_co_mm_btl import cummulative_co_mm_btl
from .cummulative_studentsmarks import cummulative_studentmarks
from .Component_calculation import Component_calculation
from .write_course_level_attainment import write_course_level_attainment
from .printout import printout
import os
import uuid
from openpyxl import load_workbook

def driver_part2(input_dir_path, output_dir_path):
    #create openpyxl workbook
    wbwrite = Workbook()
    wbwrite.remove(wbwrite.active)

    for file in os.listdir(input_dir_path):
        if file.endswith(".xlsx") and not file.startswith("Combined"):
            file_path = os.path.join(input_dir_path, file)
            wbread = load_workbook(file_path)

            wsread_input_details = None
            data={}
            Component_Details = {}
            for sheet_name in wbread.sheetnames:
                if sheet_name.endswith("Input_Details"):
                    wsread_input_details = wbread[sheet_name]
                    #create a dictionary from A2 to B11
                    for key, value in wsread_input_details.iter_rows(min_row=2, max_row=11, min_col=1, max_col=2, values_only=True):
                        data[key] = value

                    #also from A14 to B19
                    for key, value in wsread_input_details.iter_rows(min_row=14, max_row=19, min_col=1, max_col=2, values_only=True):
                        data[key] = value
                    print(data)

                     


                                         


            #copy all sheets to the new workbook
            for sheet in wbread.sheetnames:
                wswrite = wbwrite.create_sheet(sheet)
                wsread = wbread[sheet]
                for row in wsread.iter_rows(min_row=1, max_row=wsread.max_row, min_col=1, max_col=wsread.max_column, values_only=True):
                    wswrite.append(row)







            wbread.close()
    #save the workbook
    unique_id = str(uuid.uuid4())
    excel_file_name = f"Combined_{unique_id}.xlsx"
    wbwrite.save(os.path.join(output_dir_path, excel_file_name))

if __name__ == "__main__":
    input_dir_path="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_2"
    output_dir_path="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_2"
    driver_part2(input_dir_path, output_dir_path)