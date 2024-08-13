import pandas as pd
import openpyxl as op
from openpyxl.utils.cell import column_index_from_string, get_column_letter
import json
import numpy as np

from dataclasses import dataclass

def column_string_to_index(letter: str):
    return column_index_from_string(letter) - 1
def column_index_to_string(index: int):
    return get_column_letter(index+1)

@dataclass
class SheetInfo:
    name: str
    enabled: bool
    input_fp: str
    input_st: str
    row_offset: int
    columns: list[str]
    range_mode: json
    output_fp: str
    output_st: str
    mode: str
    add_name: bool

def parse_json(filepath: str) -> list[SheetInfo]:
    infos = []
    # read the json file
    with open(filepath, 'r') as project_file:
        project_json = json.load(project_file)
        # loop through every entry in the json
        for item in project_json:
            # extract the data and save it into a the sheet info array
            sheet_temp = SheetInfo(
                name=item['name'],
                enabled=item['enabled'],
                input_fp=item['input']['filepath'],
                input_st=item['input']['sheetname'],
                row_offset=item['input']['row_offset'],
                columns=item['input']['columns'],
                range_mode=item['input']['range_mode'],
                output_fp=item['output']['filepath'],
                output_st=item['output']['sheetname'],
                mode=item['output']['insert_mode'],
                add_name=item['output']['include_name'])
            infos.append(sheet_temp)
    return infos


def find_range(sheet, info):
    if info.range_mode["type"] == "End of column":
        # for each column, iterate backwards from the end of column until you find a value. Return the longest column
        column_length = 1
        for col in info.columns:
            col_index = column_string_to_index(col)
            y = sheet.shape[0]-1
            while y >= 0 and sheet.iloc[y, col_index] == None:
                y -= 1
            column_length = max(column_length, y+1)
            if column_length == sheet.shape[0]:
                break
        return column_length
    
    elif info.range_mode["type"] == "Up to row":
        return info.range_mode["row"]
    
    elif info.range_mode["type"] == "Up to code":
        # iterate until you find the code in the column or send the maximum length
        index = 0
        col_index = column_string_to_index(info.range_mode["column"])
        while index < sheet.shape[0] and sheet.iloc[index, col_index] != info.range_mode["code"]:
            index += 1
        return min(index+1, sheet.shape[0])


def next_empty_column(info: SheetInfo):
    try:
        input_st = pd.read_excel(info.input_fp, sheet_name=info.input_st, header=None)
    except Exception as e:
        return "Failed to open the input file: " + str(e)
    
    try:
        output_st = pd.read_excel(info.output_fp, sheet_name=info.output_st, header=None)
    except Exception as e:
        return "Failed to open the output file: " + str(e)

    next_column = output_st.shape[1]

    for col in info.columns:
        col_index = column_string_to_index(col)
        if col_index >= input_st.shape[1]:
            return "Columns provided aren't within the range in the input file"

    if info.range_mode["type"] == "Up to row":
        if info.range_mode["row"] >= input_st.shape[0]:
            return "Row provided isn't within range"

    # find the longest column in the input file
    column_length = find_range(input_st, info)
    
    # reshape the output excel spreadsheet so that it can fit the import
    for i in range(len(info.columns) + (1 if info.add_name else 0)):
        output_st[f'Unnamed: {next_column+1+i}'] = ' '
    for i in range(0, column_length - output_st.shape[0]):
        output_st.loc[output_st.shape[0]] = ' '

    # add the name of the entry to each row
    if info.add_name:
        for i in range(0, column_length):
            output_st.iloc[i, next_column] = info.name
        next_column += 1
    
    # add the column
    for col in info.columns:
        col_index = column_string_to_index(col)
        for i in range(column_length - info.row_offset):
            output_st.iloc[i, next_column] = input_st.iloc[i+info.row_offset, col_index]
        next_column += 1
    
    # save the result
    try:
        output_st.to_excel(info.output_fp, sheet_name=info.output_st, header=False, index=False)
    except Exception as e:
        return "Failed to write to the output file: " + str(e)


def next_empty_row(info: SheetInfo):
    try:
        input_st = pd.read_excel(info.input_fp, sheet_name=info.input_st, header=None)
    except Exception as e:
        return "Failed to open the input file: " + str(e)
    
    try:
        output_st = pd.read_excel(info.output_fp, sheet_name=info.output_st, header=None)
    except Exception as e:
        return "Failed to open the output file: " + str(e)

    next_row = output_st.shape[0]

    for col in info.columns:
        col_index = column_string_to_index(col)
        if col_index >= input_st.shape[1]:
            return "Columns provided aren't within the range in the input file"

    if info.range_mode["type"] == "Up to row":
        if info.range_mode["row"] >= input_st.shape[0]:
            return "Row provided isn't within range"

    # find the longest column in the input file
    column_length = find_range(input_st, info)
 
    # reshape the output excel spreadsheet
    for i in range(output_st.shape[1], len(info.columns) + (1 if info.add_name else 0)):
        output_st[f'Unnamed: {i}'] = ' '
    for i in range(next_row, column_length + next_row - info.row_offset):
        output_st.loc[output_st.shape[0]] = ' '

    if info.add_name:
        for i in range(next_row, next_row + column_length - info.row_offset):
            output_st.iloc[i, 0] = info.name

    next_column = 1
    for col in info.columns:
        col_index = column_string_to_index(col)
        for i in range(next_row, next_row + column_length - info.row_offset):
            output_st.iloc[i, next_column] = input_st.iloc[i-next_row+info.row_offset, col_index]
        next_column += 1
    
    try:
        output_st.to_excel(info.output_fp, sheet_name=info.output_st, header=False, index=False)
    except Exception as e:
        return "Failed to write to the output file: " + str(e)


def run():
    infos = parse_json("project_info.json")
    status = []
    for info in infos:
        if info.enabled == False:
            continue

        info.row_offset -= 1

        output = ""
        if info.mode == "Next empty column":
            output = next_empty_column(info)
        elif info.mode == "Next empty row":
            output = next_empty_row(info)
        
        if output:
            status.append(f"[error] Failed to pull columns from entry: {info.name}. {output}")
        else:
            status.append(f"[valid] Succeeded to pull columns from entry: {info.name}")

    return status