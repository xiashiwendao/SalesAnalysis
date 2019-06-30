import os
import pandas as pd


# export all the sheets to a single csv file.
# params:
#   input_excel_path : excel file path.
#   target_dir : It will be the same as input_excel_path when no content inputted.
def generate_csv(input_excel_path, target_dir):
    input_dir, input_filename = os.path.split(input_excel_path)
    output_filename, extension = os.path.splitext(input_filename)
    # path format , the file will export to the same dir as input.
    if target_dir is None or target_dir.strip() == "":
        target_dir = input_dir
    else:
        pass
    print("source_excel_path :" + input_excel_path)
    print("target_dir :" + target_dir)

    source_excel_pd = pd.read_excel(input_excel_path, sheet_name='Each', header=5, usecols="C:AB")

    source_excel_pd.to_csv(
        target_dir + "/" + output_filename + "_" + ".csv", encoding="utf-8")


generate_csv(
    r"D:\study\private_note\Einstein\NC01 NC Offtake by Product Line by SKU by Outlet Monthly Report    (2).xlsx", "")
