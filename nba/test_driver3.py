from Part_3.driver import driver_part3

import os
if __name__ == "__main__":
    input_dir_path="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_3\\TestFiles"
    output_dir_path="C:\\Users\\raman\\OneDrive - Amrita vishwa vidyapeetham\\ASE\\Projects\\NBA\\NBA_v3\\dev_19.1\\flux\\nba\\Part_3\\TestFiles"
    file_name=driver_part3(input_dir_path, output_dir_path)
    #open the file
    os.startfile(os.path.join(output_dir_path, file_name))