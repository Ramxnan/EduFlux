from Part_1.driver import driver_part1

import os


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
        "Branch":"MEE",                                                                          #set branch
        "Subject_Name":"PCE",
        "Subject_Code":"19MEE444",
        "Section":"A",
        "Semester":"Even",
        "Number_of_Students":47,
        "Number_of_COs":8}
    
    # Component_Details={"P1_I":{"Number_of_questions":3},
    #                     "P2_I":{"Number_of_questions":6},
    #                     "CA_I":{"Number_of_questions":6},
    #                     "EndSem_E":{"Number_of_questions":9}}

    # Component_Details={"P1_I":3,
    #                     "P2_I":6,
    #                     "CA_I":6,
    #                     "EndSem_E":9}
    
    # Component_Details={"P1_I":7,
    #                    "EndSem_E":13}
    Component_Details={"P1_I":7,
                        "CA_I":4,
                        "EndSem_E":13}

    # Component_Details={"P1_I":{"Number_of_questions":3},
    #                     "EndSem_E":{"Number_of_questions":3}}
    #main1(data,Component_Details)
    driver_part1(data,Component_Details, os.getcwd())