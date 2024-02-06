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
    excel_files=[]
    for file in os.listdir(input_dir_path):
        if file.endswith(".xlsx") and not file.startswith("Combined"):
            excel_files.append(file)
    excel_files.sort(reverse=True)

    for file in excel_files:
        if file.endswith(".xlsx") and not file.startswith("Combined"):
            file_path = os.path.join(input_dir_path, file)
            wbread = load_workbook(file_path, data_only=True)

            alldata={}
            Component_Details = {}
            total_students = 0
            for sheet_name in wbread.sheetnames:
                if sheet_name.endswith("Input_Details"):
                    input_details_title = sheet_name

            wsread_input_details = wbread[input_details_title]
            #create a dictionary from A2 to B11
            for key, value in wsread_input_details.iter_rows(min_row=2, max_row=11, min_col=1, max_col=2, values_only=True):
                alldata[key] = value
    
            for key, value in wsread_input_details.iter_rows(min_row=14, max_row=19, min_col=1, max_col=2, values_only=True):
                alldata[key] = value

            total_students+=alldata['Number_of_Students']

            #get only Teacher Academic_year, Batch, Branch, Subject_Name, Subject_Code, Section, Semester, Number_of_Students, Number_of_COs from alldata
            data = {key: alldata[key] for key in alldata.keys() & {'Teacher', 'Academic_year', 'Batch', 'Branch', 'Subject_Name', 'Subject_Code', 'Section', 'Semester', 'Number_of_Students', 'Number_of_COs'}}




            #extract table called Component_Details and store it in a dictionary
            table_range = wsread_input_details.tables[f'{data["Section"]}_Component_Details'].ref
            for row in wsread_input_details[table_range][1:]:
                Component_Details[row[0].value] = row[1].value

        
            

            wbwrite.create_sheet(f"{data['Section']}_Input_Details")
            wswrite = wbwrite[f"{data['Section']}_Input_Details"]
            wswrite = input_detail(data,Component_Details,wswrite)
            wswrite = indirect_co_assessment(data,wswrite)
            adjust_width(wswrite)
            wswrite = CO_PO_Table(data,wswrite)

            for key in Component_Details.keys():
                wbwrite.create_sheet(key)
                wswrite = wbwrite[key]
                wswrite.title = key
                wswrite = qn_co_mm_btl(data, key, Component_Details[key], wswrite)
                wswrite = studentmarks(data, key, Component_Details[key], wswrite)

                wswrite = cummulative_co_mm_btl(data, key, Component_Details[key], wswrite)   
                wswrite = cummulative_studentmarks(data, key, Component_Details[key], wswrite)

            wbwrite.create_sheet(f"{data['Section']}_Internal_Components")
            wswrite = wbwrite[f"{data['Section']}_Internal_Components"]
            wswrite = Component_calculation(data,Component_Details,wswrite,"I")

            wbwrite.create_sheet(f"{data['Section']}_External_Components")
            wswrite = wbwrite[f"{data['Section']}_External_Components"]
            wswrite = Component_calculation(data,Component_Details,wswrite,"E")

            wbwrite.create_sheet(f"{data['Section']}_Course_level_Attainment")
            wswrite = wbwrite[f"{data['Section']}_Course_level_Attainment"]
            wswrite=write_course_level_attainment(data, Component_Details, wswrite)

            wbwrite.create_sheet(f"{data['Section']}_Printout")
            wswrite = wbwrite[f"{data['Section']}_Printout"]
            wswrite=printout(wswrite,data)

            #copy data from all the sheets of wbread to wbwrite
            for sheet in wbread.sheetnames:
                wsread = wbread[sheet]
                wswrite = wbwrite[sheet]
                for row in wsread.iter_rows(min_row=1, max_row=wsread.max_row, min_col=1, max_col=wsread.max_column):
                    for cell in row:
                        #if error occurs while copying the cell, skip the cell
                        try:
                            wswrite[cell.coordinate].value = cell.value
                        except:
                            pass









            wbread.close()
    #save the workbook
    unique_id = str(uuid.uuid4())
    excel_file_name = f"Combined_{unique_id}.xlsx"
    wbwrite.save(os.path.join(output_dir_path, excel_file_name))

if __name__ == "__main__":
    input_dir_path="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_2"
    output_dir_path="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_2"
    driver_part2(input_dir_path, output_dir_path)