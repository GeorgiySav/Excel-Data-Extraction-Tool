import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import backups.program as program

entries = []

def load_from_json():
    global entries
    with open('project_info.json', 'r') as json_file:
        entries = json.load(json_file)
        for entry in entries:
            my_tag = True if entry["enabled"] else False
            table.insert('', index='end', values=(
                entry["name"], 
                entry["input_fp"], 
                entry["input_st"], 
                entry["output_fp"], 
                entry["output_st"],
                entry["row_offset"],
                entry["columns"],
                entry["enabled"],
                entry["mode"],
                entry["add_name"],
                entry["column_offset"]),
                tags=(my_tag))
    json_file.close()
    print(entries)

def add_column():
    column = column_entry.get().strip().upper()
    if column.isalpha() == False:
        return

    frame = tk.Frame(columns_frame, width=125, height=15)
    frame.pack_propagate(False)
    label = tk.Label(frame, text=column)
    label.pack(side=tk.LEFT)
    remove_button = tk.Button(frame, text='-', command=lambda: remove_column(frame))
    remove_button.pack(side=tk.RIGHT)
    frame.pack(fill=tk.X)
    column_entry.delete(0, tk.END)

def remove_column(frame):
    frame.destroy()

def delete_from_json():
    target_name = name_var.get()
    for i in range(len(entries)):
        if entries[i]["name"] == target_name:
            del entries[i]
            break
    for item_id in table.get_children():
        item = table.item(item_id)
        if item['values'][0] == target_name:
            table.delete(item_id)
            return

def update_json():
    target_name = name_var.get()
    i = 0

    while i < len(entries):
        if entries[i]["name"] == target_name:
            entries[i]["input_fp"] = input_fp_var.get()
            entries[i]["input_st"] = input_st_var.get()
            entries[i]["output_fp"] = output_fp_var.get()
            entries[i]["output_st"] = output_st_var.get()
            entries[i]["row_offset"] = int(row_offset_var.get())
            entries[i]["columns"] = [frame.winfo_children()[0].cget("text") for frame in columns_frame.winfo_children()]
            entries[i]["enabled"] = enabled_var.get()
            mode = mode_cb.()
            break
        i += 1

    for item_id in table.get_children():
        item = table.item(item_id)
        if item['values'][0] == target_name:
            my_tag = True if entries[i]["enabled"] else False  
            table.item(item_id, values=(
                entries[i]["name"], 
                entries[i]["input_fp"], 
                entries[i]["input_st"], 
                entries[i]["output_fp"], 
                entries[i]["output_st"],
                entries[i]["row_offset"],
                entries[i]["columns"],
                entries[i]["enabled"],
                entries[i]["mode"],
                entries[i]["add_name"],
                entries[i]["column_offset"]),
                tags=(my_tag))
            return
            

def save_to_json():
    # check if there already exists a file with the same name
    for entry in entries:
        if entry["name"] == name_var.get():
            return

    columns = [frame.winfo_children()[0].cget("text") for frame in columns_frame.winfo_children()]
    data = {
        "name": name_var.get(),
        "input_fp": input_fp_var.get(),
        "input_st": input_st_var.get(),
        "output_fp": output_fp_var.get(),
        "output_st": output_st_var.get(),
        "row_offset": int(row_offset_var.get()), 
        "columns": columns,
        "enabled": enabled_var.get()
    }
    
    # Add entry to the table
    entries.append(data)
    my_tag = True if data["enabled"] else False
    table.insert('', index='end', values=(
        data["name"], 
        data["input_fp"], 
        data["input_st"], 
        data["output_fp"], 
        data["output_st"],
        data["row_offset"],
        ','.join([x for x in data["columns"]]),
        data["enabled"]),
        tags=(my_tag))
    #print(entries)
    clear_form()

def save_to_file():
    with open('project_info.json', 'w') as json_file:
        json.dump(entries, json_file, indent=2)

def clear_form():
    name_var.set("")
    input_fp_var.set("")
    input_st_var.set("")
    output_fp_var.set("")
    output_st_var.set("")
    row_offset_var.set("")
    enabled_var.set(False)
    for frame in columns_frame.winfo_children():
        frame.destroy()

def browse_file(var):
    filename = filedialog.askopenfilename()
    var.set(filename)

def on_table_select(event):
    selected_item = table.focus()

    if selected_item:
        print(table.item(selected_item)['values'])

        target_name = table.item(selected_item)['values'][0]
        print(target_name)
        print(entries)
        
        for entry in entries:
            print(entry)
            if target_name == entry["name"]:
                selected_data = entry
        
        name_var.set(selected_data["name"])
        input_fp_var.set(selected_data["input_fp"])
        input_st_var.set(selected_data["input_st"])
        output_fp_var.set(selected_data["output_fp"])
        output_st_var.set(selected_data["output_st"])
        row_offset_var.set(selected_data["row_offset"])
        enabled_var.set(selected_data["enabled"])
        for frame in columns_frame.winfo_children():
            frame.destroy()
        for column in selected_data["columns"]:
            column_entry.delete(0, tk.END)
            column_entry.insert(0, column)
            add_column()

# Create the main window
root = tk.Tk()
root.title("Excel Data Extraction Form")

# Define Tkinter variables
name_var = tk.StringVar()
input_fp_var = tk.StringVar()
input_st_var = tk.StringVar()
output_fp_var = tk.StringVar()
output_st_var = tk.StringVar()
row_offset_var = tk.StringVar()
enabled_var = tk.BooleanVar()
add_name_var = tk.BooleanVar()
column_offset_var = tk.StringVar()

# Create and place the widgets
frame = tk.Frame(root)
frame.pack()

file_info_frame = tk.LabelFrame(frame, text="")
file_info_frame.grid(row=0, column=0, padx=5, pady=(10, 0))

name_frame = tk.LabelFrame(file_info_frame, text="File ID and Status")
name_frame.grid(row=0, column=0, padx=10, pady=5, sticky="news")

tk.Label(name_frame, text="Name of Project").grid(row=0, column=0, sticky="ew")
tk.Entry(name_frame, textvariable=name_var).grid(row=0, column=1, sticky="ew")

tk.Checkbutton(name_frame, text="Enabled", variable=enabled_var, onvalue=True, offvalue=False, anchor="w").grid(row=1, column=0, sticky="news")

for widget in name_frame.winfo_children():
    widget.grid_configure(padx=5, pady=3)

input_excel_frame = tk.LabelFrame(file_info_frame, text="Input Excel Workbook Information")
input_excel_frame.grid(row=0, column=1, rowspan=2, padx=10, pady=5, sticky="nsew")

tk.Label(input_excel_frame, text="Filepath", anchor="w").grid(row=0, column=0, sticky="ew")
tk.Entry(input_excel_frame, textvariable=input_fp_var).grid(row=0, column=1, sticky="ew")
tk.Button(input_excel_frame, text="Browse", command=lambda: browse_file(input_fp_var)).grid(row=0, column=2, sticky="ew")

tk.Label(input_excel_frame, text="Sheet Name", anchor="w").grid(row=1, column=0, sticky="ew")
tk.Entry(input_excel_frame, textvariable=input_st_var).grid(row=1, column=1, columnspan=2, sticky="ew")

tk.Label(input_excel_frame, text="Offset Row", anchor="w").grid(row=2, column=0, sticky="ew")
tk.Entry(input_excel_frame, textvariable=row_offset_var).grid(row=2, column=1, columnspan=2, sticky="ew")

tk.Label(input_excel_frame, text="Columns", anchor="w").grid(row=3, column=0, sticky="ew")
column_entry = tk.Entry(input_excel_frame)
column_entry.grid(row=3, column=1, sticky="ew")
tk.Button(input_excel_frame, text="Add", command=add_column).grid(row=3, column=2, sticky="ew")

columns_frame = tk.Frame(input_excel_frame)
columns_frame.grid(row=4, column=1, sticky="ew")

for widget in input_excel_frame.winfo_children():
    widget.grid_configure(padx=5, pady=3)

output_excel_frame = tk.LabelFrame(file_info_frame, text="Output Excel Workbook Information")
output_excel_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

tk.Label(output_excel_frame, text="Filepath", anchor="w").grid(row=0, column=0, sticky="ew")
tk.Entry(output_excel_frame, textvariable=output_fp_var).grid(row=0, column=1, sticky="ew")
tk.Button(output_excel_frame, text="Browse", command=lambda: browse_file(output_fp_var)).grid(row=0, column=2, sticky="ew")

tk.Label(output_excel_frame, text="Sheet Name", anchor="w").grid(row=1, column=0, sticky="ew")
tk.Entry(output_excel_frame, textvariable=output_st_var).grid(row=1, column=1, columnspan=2, sticky="ew")

ttk.Separator(output_excel_frame, orient="horizontal").grid(row=2, column=0, columnspan=3, sticky="ew")

tk.Label(output_excel_frame, text="Mode", anchor="w").grid(row=3, column=0, sticky="ew")
mode_cb = ttk.Combobox(output_excel_frame, state="readonly", values=["Next empty column", "Next empty row"]).grid(row=3, column=1, columnspan=2, sticky="ew")

tk.Label(output_excel_frame, text="Column Offset", anchor="w").grid(row=4, column=0, sticky="ew")
tk.Entry(output_excel_frame, textvariable=column_offset_var).grid(row=4, column=1, columnspan=2, sticky="ew")

tk.Checkbutton(output_excel_frame, text="Include Name", variable=add_name_var).grid(row=5, column=0)

for widget in output_excel_frame.winfo_children():
    widget.grid_configure(padx=5, pady=3)

submit_buttons_frame = tk.Frame(file_info_frame)
submit_buttons_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=(0,5), sticky="news")

tk.Button(submit_buttons_frame, text="Delete from Table", command=delete_from_json).pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)
tk.Button(submit_buttons_frame, text="Update entry in Table", command=update_json).pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)
tk.Button(submit_buttons_frame, text="Save to Table", command=save_to_json).pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)

# Create the table to display saved entries
table_frame = tk.LabelFrame(frame)
table_frame.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

table = ttk.Treeview(table_frame, columns=(1, 2, 3, 4, 5, 6, 7, 8), show="headings", height="16")
#table.grid(row=1, column=0, padx=10, pady=10, sticky="news")
table.pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)

table.heading(1, text="Name", anchor="center")
table.heading(2, text="Input FP", anchor="center")
table.heading(3, text="Input SN", anchor="center")
table.heading(4, text="Output FP", anchor="center")
table.heading(5, text="Output SN", anchor="center")
table.heading(6, text="Offset", anchor="center")
table.heading(7, text="Columns", anchor="center")
table.heading(8, text="Enabled", anchor="center")

table.column("#1", anchor="w", stretch=True, width=70)
table.column("#2", anchor="w", stretch=True, width=80)
table.column("#3", anchor="w", stretch=True, width=60)
table.column("#4", anchor="w", stretch=True, width=80)
table.column("#5", anchor="w", stretch=True, width=60)
table.column("#6", anchor="w", stretch=True, width=50)
table.column("#7", anchor="w", stretch=True, width=75)
table.column("#8", anchor="w", stretch=True, width=50)

table.tag_configure(False, background="firebrick1")
table.tag_configure(True, background="lightgreen")

table_scroll_bar = tk.Scrollbar(table_frame, orient="vertical", command=table.yview)
table.configure(yscrollcommand=table_scroll_bar.set)
#table_scroll_bar.grid(row=1, column=1, sticky="ns")
table_scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)

table.bind('<<TreeviewSelect>>', on_table_select)

tk.Button(root, text="Save To File", command=save_to_file).pack(padx=5, pady=(0,5), side=tk.LEFT, fill=tk.X, expand=tk.YES)

tk.Button(root, text="Pull Data", command=lambda: program.run()).pack(padx=5, pady=(0,5), side=tk.LEFT, fill=tk.X, expand=tk.YES)

load_from_json()

# Run the application
root.mainloop()
