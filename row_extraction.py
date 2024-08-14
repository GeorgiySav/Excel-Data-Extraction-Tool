import pandas as pd
import openpyxl as op
from openpyxl.utils.cell import column_index_from_string, get_column_letter
import json
import numpy as np

from dataclasses import dataclass

# for converting a column to integer. e.g. A to 0
def column_string_to_index(letter: str):
    return column_index_from_string(letter) - 1
# for converting an integer to column. e.g. 0 to A
def column_index_to_string(index: int):
    return get_column_letter(index+1)


@dataclass
class SheetInfo:
    name: str
    enabled: bool
    input_fp: str
    input_st: str
    selection_mode: json
    output_fp: str
    output_st: str
    add_name: bool


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
                    selection_mode=item['input']['selection_mode'],
                    output_fp=item['output']['filepath'],
                    output_st=item['output']['sheetname'],
                    add_name=item['output']['include_name'])
                infos.append(sheet_temp)#
            except Exception as e:
                raise e
    return infos


def pull_rows(info: SheetInfo):
    try:
        input_st = pd.read_excel(info.input_fp, sheet_name=info.input_st, header=None)
    except Exception as e:
        return "Failed to open the input file: " + str(e)
    
    try:
        output_st = pd.read_excel(info.output_fp, sheet_name=info.output_st, header=None)
    except Exception as e:
        return "Failed to open the output file: " + str(e)

    # find the row
    row = []
    if info.selection_mode["type"] == "row":
        try:
            row = input_st.iloc[[info.selection_mode["row"]-1]].values.flatten().tolist()
        except Exception as e:
            return "Provided row out of range: " + str(e)
    elif info.selection_mode["type"] == "code":
        # iterate through the column until you find the code
        # if the code isn't there, then the last row will be extracted
        index = 0
        try:
            col_index = column_string_to_index(info.selection_mode["column"])
        except Exception as e:
            return "Provided column out of range: " + str(e)
        while index < input_st.shape[0] and input_st.iloc[index, col_index] != info.selection_mode["code"]:
            index += 1
        index = min(index, input_st.shape[0] - 1)
        row = input_st.iloc[[index]].values.flatten().tolist()

    # add the name at the start of the row if needed
    if info.add_name:
        row = [info.name] + row

    # add the row to the next row available
    output_st = pd.concat([output_st, pd.DataFrame([row])], ignore_index=True)

    try:
        output_st.to_excel(info.output_fp, sheet_name=info.output_st, header=False, index=False)
    except Exception as e:
        return "Failed to write to the output file: " + str(e)
    

def run():
    try:
        infos = parse_json("row_project_info.json")
    except Exception as e:
        return [f"Failed to open the parse files: {str(e)}"]
    
    status = []
    for info in infos:
        if info.enabled == False:
            continue
        
        output = pull_rows(info)
        if output:
            status.append(f"[error] Failed to pull row from entry: {info.name}. {output}")
        else:
            status.append(f"[valid] Succeeded to pull row from entry: {info.name}")
    return status