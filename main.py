import os

from excelbulk.ExcelBulk import ExcelBulk
import argparse

# Construct an argument parser
all_args = argparse.ArgumentParser()

# Add arguments to the parser
all_args.add_argument("-i", "--InFolder",type=str, required=True,
   help="The Folder to Read")
all_args.add_argument("-o", "--OutFolder",type=str, required=True,
   help="The folder to save the")
args = vars(all_args.parse_args())

if '__main__' == __name__:
    # Variables and Paths
    current_dir = os.path.abspath(os.getcwd())
    print(f"This is your Current Directory: \n {current_dir}")
    print(f"These are the current folders/files in this directory: \n {os.listdir(current_dir)}")

    #read_folder_name = input(str("Please enter the folder you want to read: "))
    read_folder_name = str(args['InFolder'])
    print(f"\nFolder Name:\t {read_folder_name}")

    read_path = os.path.join(current_dir,read_folder_name)

    # out_folder = input(str("Please enter the folder you want to output your results to: "))
    out_folder = str(args['OutFolder'])
    out_path = os.path.join(current_dir,out_folder)

    if not os.path.exists(out_path):
        os.makedirs(out_path)
    print(f"This is your reading path: {read_path}")
    print(f"This is your output  path: {out_path}")

    Excel = ExcelBulk(current_dir,read_folder_name,out_folder)
    # Main execution
    raw_files = Excel.get_excel_list()

    for n in range(0,len(raw_files)):
        sheet_names = Excel.get_sheet_names(raw_files,n)

    #Excel.change_sheet_name(raw_files,'Alpha','alpha')
    Excel.sort_sheets_in_folder_alphabetically(raw_files)

    sheet_to_remove = 'Subject Info'
    step_1_folder, modified_path = Excel.remove_sheet(raw_files,sheet_to_remove)
    os.chdir(out_folder)
    #
    for sn in range(0,len(sheet_names)):
         Excel.multi_file_sheet_to_excel(step_1_folder,raw_files,sn,f"{sheet_names[sn]}.xlsx")

    # runs the csv_from_excel function:
    Excel.csv_from_excel()

    num_rows = 20
    Excel.copy_data_and_split(num_rows)

    print("DATA PROCESSING PROCESS DONE")















