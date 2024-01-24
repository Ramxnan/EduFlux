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

def main1(data,Component_Details, file_path):
    #create openpyxl workbook
    wb = Workbook()
    wb.remove(wb.active)


    wb.create_sheet("Input_Details")
    ws = wb["Input_Details"]
    ws = input_detail(data,ws)
    ws = indirect_co_assessment(data,ws)
    adjust_width(ws)
    ws = CO_PO_Table(data,ws)



    
    #get number of components
    number_of_components = len(Component_Details.keys())

    #iterate throught Keys of Component_Details and make a worksheet for each key
    for key in Component_Details.keys():
        wb.create_sheet(key)
        ws = wb[key]
        ws.title = key
        ws = qn_co_mm_btl(data, key, Component_Details[key], ws)
        ws = studentmarks(data, key, Component_Details[key], ws)

        ws = cummulative_co_mm_btl(data, key, Component_Details[key], ws)   
        ws = cummulative_studentmarks(data, key, Component_Details[key], ws)


        #adjust_width(ws)
    
    wb.create_sheet("Internal_Components")
    ws = wb["Internal_Components"]
    ws = Component_calculation(data,Component_Details,ws,"I")

    wb.create_sheet("External_Components")
    ws = wb["External_Components"]
    ws = Component_calculation(data,Component_Details,ws,"E")

    wb.create_sheet("Course_level_Attainment")
    ws = wb["Course_level_Attainment"]
    ws=write_course_level_attainment(data, Component_Details, ws)

    wb.create_sheet("Printout")
    ws = wb["Printout"]
    ws=printout(ws,data)

    #save workbook
    #wb.save(f"{data['Batch']}_{data['Subject_Code']}_{data['Subject_Name']}.xlsx")
    excel_file_name = f"{data['Batch']}_{data['Subject_Code']}_{data['Subject_Name']}_{data['Section']}_{data['Semester']}.xlsx"
    excel_file_name.replace(" ","_")
    full_path = os.path.join(file_path, excel_file_name)
    wb.save(full_path)
    return excel_file_name    


if __name__ == "__main__":
    
    # data={
    #     "Teacher":"Dr. S. S. Patil",                                                              #set teacher name
    #     "Academic_year":"2022-2023",                                                              #set academic year
    #     "Semester":7,                                                                                 #set semester
    #     "Branch":"CSE",                                                                          #set branch
    #     "Batch":2019,                                                                             #set batch
    #     "Section":"A",                                                                           #set section
    #     "Subject_Code":"19CSE345",                                                            #set subject code
    #     "Subject_Name":"Computer system and architecture",                          #set subject name
    #     "Number_of_Students":47,
    #     "Number_of_COs":4,
    #     "Internal":50,
    #     "Direct":80,
    #     "Default threshold %":70,
    #      "target":60
    #         }
    
    data={
        "Teacher":"Dr. S. S. Patil",                                                              #set teacher name
        "Academic_year":"2022-2023",  
        "Batch":2019,
        "Branch":"CSE",                                                                          #set branch
        "Subject_Name":"Computer system and architecture",
        "Subject_Code":"19CSE345",
        "Section":"A",
        "Semester":"Even",
        "Number_of_Students":47,
        "Number_of_COs":4}
    
    data={
        "Teacher":"Dr. S. S. Patil",                                                              #set teacher name
        "Academic_year":"2022-2023",  
        "Batch":2019,
        "Branch":"CSE",                                                                          #set branch
        "Subject_Name":"PCE",
        "Subject_Code":"19MEE444",
        "Section":"A",
        "Semester":"Even",
        "Number_of_Students":10,
        "Number_of_COs":4}
    
    # Component_Details={"P1_I":{"Number_of_questions":3},
    #                     "P2_I":{"Number_of_questions":6},
    #                     "CA_I":{"Number_of_questions":6},
    #                     "EndSem_E":{"Number_of_questions":9}}

    # Component_Details={"P1_I":3,
    #                     "P2_I":6,
    #                     "CA_I":6,
    #                     "EndSem_E":9}
    
    Component_Details={"P1_I":7,
                       "EndSem_E":13}
    # Component_Details={"P1_I":3,
    #                     "CA_I":6,
    #                     "EndSem_E":9}

    # Component_Details={"P1_I":{"Number_of_questions":3},
    #                     "EndSem_E":{"Number_of_questions":3}}
    #main1(data,Component_Details)
    main1(data,Component_Details, os.getcwd())