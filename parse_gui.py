from pygui_utils import *
import json

import parse_extraction

def table_row_callback(sender, app_data, user_data):
    name = user_data["name"]
    enabled = user_data["enabled"]

    input_fp = user_data["input"]["filepath"]
    input_st = user_data["input"]["sheetname"]
    column = user_data["input"]["column"]

    output_fp = user_data["output"]["filepath"]
    
    dpg.set_value("Parse Name Input", name)
    dpg.set_value("Parse Enabled Input", enabled)

    dpg.set_value("Parse Input FP Input", input_fp)
    dpg.set_value("Parse Input ST Input", input_st)
    dpg.set_value("Parse Column Input", column)

    dpg.set_value("Parse Output FP Input", output_fp)


def save_to_parse_table():
    name =  dpg.get_value("Parse Name Input")
    if name == "":
        return
    rows = dpg.get_item_children("parse table", 1)
    for row in rows:
        cells = dpg.get_item_children(row, 1)
        if dpg.get_item_label(cells[0]) == name:
            return
        
    enabled = dpg.get_value("Parse Enabled Input")

    input_fp = dpg.get_value("Parse Input FP Input")
    input_st = dpg.get_value("Parse Input ST Input")
    column = dpg.get_value("Parse Column Input")

    output_fp = dpg.get_value("Parse Output FP Input")

    entry = {
        "name": name,
        "enabled": enabled,
        "input": {
            "filepath": input_fp,
            "sheetname": input_st,
            "column": column
        },
        "output": {
            "filepath": output_fp,
        }
    }
    print(entry)

    with dpg.table_row(parent="parse table"):
        dpg.add_selectable(label=name, span_columns=True, callback=table_row_callback, user_data=entry)
        dpg.add_selectable(label=enabled, span_columns=True, callback=table_row_callback, user_data=entry)

        dpg.add_selectable(label=input_fp, span_columns=True, callback=table_row_callback, user_data=entry)
        dpg.add_selectable(label=input_st, span_columns=True, callback=table_row_callback, user_data=entry)
        dpg.add_selectable(label=column, span_columns=True, callback=table_row_callback, user_data=entry)

        dpg.add_selectable(label=output_fp, span_columns=True, callback=table_row_callback, user_data=entry)


def update_parse_entry():
    name = dpg.get_value("Parse Name Input")
    if name == "":
        return
        
    rows = dpg.get_item_children("parse table", 1)
    for row in rows:
        cells = dpg.get_item_children(row, 1)
        if dpg.get_item_label(cells[0]) != name:
            continue

        enabled = dpg.get_value("Parse Enabled Input")

        input_fp = dpg.get_value("Parse Input FP Input")
        input_st = dpg.get_value("Parse Input ST Input")
        column = dpg.get_value("Parse Column Input")

        output_fp = dpg.get_value("Parse Output FP Input")

        entry = {
            "name": name,
            "enabled": enabled,
            "input": {
                "filepath": input_fp,
                "sheetname": input_st,
                "column": column,
            },
            "output": {
                "filepath": output_fp,
            }
        }

        dpg.set_item_user_data(cells[0], entry)

        dpg.set_item_label(cells[1], enabled)
        dpg.set_item_user_data(cells[1], entry)
        
        dpg.set_item_label(cells[2], input_fp)
        dpg.set_item_user_data(cells[2], entry)

        dpg.set_item_label(cells[3], input_st)
        dpg.set_item_user_data(cells[3], entry)

        dpg.set_item_label(cells[4], column)
        dpg.set_item_user_data(cells[4], entry)

        dpg.set_item_label(cells[5], output_fp)
        dpg.set_item_user_data(cells[5], entry)


def delete_parse_entry():
    name =  dpg.get_value("Parse Name Input")

    rows = dpg.get_item_children("parse table", 1)
    for row in rows:
        cells = dpg.get_item_children(row, 1)
        if dpg.get_item_label(cells[0]) != name:
            continue
        dpg.delete_item(row)
        break


def save_to_parse_file():
    entries = []

    rows = dpg.get_item_children("parse table", 1)
    for row in rows:
        cells = dpg.get_item_children(row, 1)

        entry = {}
        entry["name"] = dpg.get_item_configuration(cells[0])["label"]
        entry["enabled"] = True if dpg.get_item_configuration(cells[1])["label"] == "True" else False
        entry["input"] = {}
        entry["input"]["filepath"] = dpg.get_item_configuration(cells[2])["label"]
        entry["input"]["sheetname"] = dpg.get_item_configuration(cells[3])["label"]
        entry["input"]["column"] = dpg.get_item_configuration(cells[4])["label"]
        entry["output"] = {}
        entry["output"]["filepath"] = dpg.get_item_configuration(cells[5])["label"]

        print(entry)
        entries.append(entry)
    
    with open('parse_project_info.json', 'w') as json_file:
        json.dump(entries, json_file, indent=2)


def add_error_msg_to_console(message):
    dpg.add_text(message, parent="console_window", color=(196, 43, 43))
    dpg.set_y_scroll("console_window", dpg.get_y_scroll_max("console_window"))


def add_valid_msg_to_console(message):
    dpg.add_text(message, parent="console_window", color=(69, 214, 69))
    dpg.set_y_scroll("console_window", dpg.get_y_scroll_max("console_window"))


def parse_file():
    stati = parse_extraction.run()
    for status in stati:
        if status[1] == "v":
            add_valid_msg_to_console(status)
        else:
            add_error_msg_to_console(status)


def load_to_parse_table(filepath: str):
    with open(filepath, 'r') as json_file:
        entries = json.load(json_file)
        print(f'entries {entries}')
        for entry in entries:
            name = entry["name"]
            enabled = entry["enabled"]

            input_fp = entry["input"]["filepath"]
            input_st = entry["input"]["sheetname"]
            column = entry["input"]["column"]

            output_fp = entry["output"]["filepath"]

            with dpg.table_row(parent="parse table"):
                dpg.add_selectable(label=name, span_columns=True, callback=table_row_callback, user_data=entry)
                dpg.add_selectable(label=enabled, span_columns=True, callback=table_row_callback, user_data=entry)

                dpg.add_selectable(label=input_fp, span_columns=True, callback=table_row_callback, user_data=entry)
                dpg.add_selectable(label=input_st, span_columns=True, callback=table_row_callback, user_data=entry)
                dpg.add_selectable(label=column, span_columns=True, callback=table_row_callback, user_data=entry)

                dpg.add_selectable(label=output_fp, span_columns=True, callback=table_row_callback, user_data=entry)

    json_file.close()


def parse_gui():
    with dpg.child_window(autosize_x=True, height=75):
        dpg.add_input_text(label="Name", tag="Parse Name Input")
        dpg.add_checkbox(label="Enabled", tag="Parse Enabled Input")
    
    with dpg.child_window(autosize_x=True, height=125):
        create_file_input(entry_label="Input Filepath", input_tag="Parse Input FP Input")
        dpg.add_input_text(label="Input Sheetname", tag="Parse Input ST Input")
        dpg.add_input_text(label="Column", tag="Parse Column Input")

    with dpg.child_window(autosize_x=True, height=75):
        create_file_input(entry_label="Output Filepath", input_tag="Parse Output FP Input")

    with dpg.group(horizontal=True):
        dpg.add_button(label="Add to table", callback=save_to_parse_table)
        dpg.add_button(label="Update entry", callback=update_parse_entry)
        dpg.add_button(label="Delete entry", callback=delete_parse_entry)
        dpg.add_button(label="Save to file", callback=save_to_parse_file)
        dpg.add_button(label="Pull data", callback=parse_file)

    with dpg.table(tag="parse table", header_row=True, row_background=True, policy=dpg.mvTable_SizingStretchSame,
                borders_innerH=True, borders_outerH=True, borders_innerV=True,
                borders_outerV=True):
        dpg.add_table_column(label="Name")
        dpg.add_table_column(label="Enabled")
        dpg.add_table_column(label="Input FP")
        dpg.add_table_column(label="Input ST")
        dpg.add_table_column(label="Column")
        dpg.add_table_column(label="Output FP")

    load_to_parse_table("parse_project_info.json")