import pandas as pd
from openpyxl.utils.cell import column_index_from_string, get_column_letter
import json

from dataclasses import dataclass


def column_string_to_index(letter: str):
    return column_index_from_string(letter) - 1
def column_index_to_string(index: int):
    return get_column_letter(index + 1)


@dataclass
class SheetInfo:
    name: str
    enabled: bool
    input_fp: str
    input_st: str
    column: str
    output_fp: str


def parse_json(filepath: str) -> list[SheetInfo]:
    infos = []
    # read the json file
    with open(filepath, 'r') as project_file:
        project_json = json.load(project_file)
        # loop through every entry in the json
        for item in project_json:
            # extract the data and save it into a the sheet info array
            try:
                sheet_temp = SheetInfo(
                    name=item['name'],
                    enabled=item['enabled'],
                    input_fp=item['input']['filepath'],
                    input_st=item['input']['sheetname'],
                    column=item['input']['column'],
                    output_fp=item['output']['filepath'])
                infos.append(sheet_temp)
            except Exception as e:
                raise e
    return infos


def parse_extraction(info: SheetInfo):
    try:
        input_st = pd.read_excel(info.input_fp, sheet_name=info.input_st, header=None)
    except Exception as e:
        return "Failed to open the input file: " + str(e)
    
    try:
        output_wk = pd.ExcelFile(info.output_fp)
    except Exception as e:
        return "Failed to open the output file: " + str(e)
    
    try:
        col_index = column_string_to_index(info.column)
    except Exception as e:
        return "Invalid column: " + str(e)

    try:
        input_st = input_st.sort_values(by=input_st.columns[col_index], na_position='first', kind="mergesort")
    except Exception as e:
        return "Could not sort input file, column is likely out of range: " + str(e)

    # skip any empty or nan values 
    index = 0
    while index < input_st.shape[0] and (input_st.iloc[index, col_index] == None or pd.isna(input_st.iloc[index, col_index])):
        index += 1

    while index < input_st.shape[0]:
        start_index = index
        start_name = input_st.iloc[start_index, col_index]
        # excel sheet names can only be at most 30 characters long
        clamped_start_name = start_name[max(0,len(start_name)-31):]

        # iterate until you find a value that is different to the beginning
        index += 1
        while index < input_st.shape[0] and input_st.iloc[index, col_index] == start_name:
            index += 1
        
        # load sheet if it exists, otherwise create a new one
        sheet = pd.DataFrame()
        if clamped_start_name in output_wk.sheet_names:
            sheet = pd.read_excel(output_wk, clamped_start_name, header=None)
        
        # append the rows to the sheet
        for i in range(start_index, index):
            row = input_st.iloc[[i]].values.flatten().tolist()
            sheet = pd.concat([sheet, pd.DataFrame([row])], ignore_index=True)
        
        # save the sheet
        try:
            with pd.ExcelWriter(info.output_fp, mode="a", if_sheet_exists="overlay") as writer:
                sheet.to_excel(writer, sheet_name=clamped_start_name, header=False, index=False)
        except Exception as e:
            return "Could not write to output: " + str(e)

def run():
    try:
        infos = parse_json("parse_project_info.json")
    except Exception as e:
        return [f"Failed to open the parse files: {str(e)}"]
    
    status = []
    for info in infos:
        if info.enabled == False:
            continue
        
        output = parse_extraction(info)
        if output:
            status.append(f"[error] Failed to parse rows in entry: {info.name}. {output}")
        else:
            status.append(f"[valid] Succeeded to parse rows in entry: {info.name}")
    return status