import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
#import program
import column_extraction

def load_from_json():
    with open('project_info.json', 'r') as json_file:
        entries = json.load(json_file)
        for entry in entries:
            enabled_var.set(entry['enabled'])
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
    for item_id in table.get_children():
        item = table.item(item_id)
        if item['values'][0] == target_name:
            table.delete(item_id)
            return

def update_table():
    target_name = name_var.get()

    for item_id in table.get_children():
        item = table.item(item_id)
        if item['values'][0] == target_name:
            my_tag = True if enabled_var.get() else False  
            table.item(item_id, values=(
                target_name, 
                input_fp_var.get(),
                input_st_var.get(),  
                output_fp_var.get(), 
                output_st_var.get(),
                int(row_offset_var.get()),
                [frame.winfo_children()[0].cget("text") for frame in columns_frame.winfo_children()],
                enabled_var.get(),
                mode_cb_var.get(),
                add_name_var.get(),
                column_offset_var.get()),
                tags=(my_tag))
            return
            
def save_to_table():
    # TODO: check if there already exists a file with the same name
    
    # Add entry to the table
    my_tag = True if enabled_var.get() else False
    table.insert('', index='end', values=(
        name_var.get(), 
        input_fp_var.get(), 
        input_st_var.get(), 
        output_fp_var.get(), 
        output_st_var.get(),
        int(row_offset_var.get()),
        ' '.join([x for x in [frame.winfo_children()[0].cget("text") for frame in columns_frame.winfo_children()]]),
        enabled_var.get(),
        mode_cb_var.get(),
        add_name_var.get(),
        column_offset_var.get()),
        tags=(my_tag)
    )


def save_to_file():
    json_data = []

    for id in table.get_children():
        record = table.item(id)
        columns = record['values'][6].split(' ')
        data = {
            "name": record['values'][0],
            "input_fp": record['values'][1],
            "input_st": record['values'][2],
            "output_fp": record['values'][3],
            "output_st": record['values'][4],
            "row_offset": record['values'][5], 
            "columns": columns,
            "enabled": True if record['values'][7] == 'True' else False,
            "mode": record['values'][8],
            "add_name": True if record['values'][9] == 'True' else False,
            "column_offset": record['values'][10]
        }
        json_data.append(data)


    with open('project_info.json', 'w') as json_file:
        json.dump(json_data, json_file, indent=2)

def browse_file(var):
    filename = filedialog.askopenfilename()
    var.set(filename)

def on_table_select(event):
    selected_item = table.focus()

    if selected_item:
        name_var.set(table.item(selected_item)['values'][0])
        input_fp_var.set(table.item(selected_item)['values'][1])
        input_st_var.set(table.item(selected_item)['values'][2])
        output_fp_var.set(table.item(selected_item)['values'][3])
        output_st_var.set(table.item(selected_item)['values'][4])
        row_offset_var.set(table.item(selected_item)['values'][5])
        for frame in columns_frame.winfo_children():
            frame.destroy()
        for column in table.item(selected_item)['values'][6]:
            column_entry.delete(0, tk.END)
            column_entry.insert(0, column)
            add_column()
        enabled_var.set(table.item(selected_item)['values'][7])
        mode_cb_var.set(table.item(selected_item)['values'][8])
        add_name_var.set(table.item(selected_item)['values'][9])
        column_offset_var.set(table.item(selected_item)['values'][10])

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
mode_cb_var = tk.StringVar()

# Create and place the widgets
frame = tk.Frame(root)
frame.pack()

file_info_frame = tk.LabelFrame(frame, text="")
#file_info_frame.grid(row=0, column=0, padx=5, pady=10, sticky="nsew")
file_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES, padx=5, pady=5)

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

ttk.Separator(input_excel_frame, orient="horizontal").grid(row=5, column=0, columnspan=3, sticky="ew")

tk.Label(input_excel_frame, text="Last row code", anchor="w").grid(row=6, column=0, sticky="ew")
tk.Entry(input_excel_frame).grid(row=6, column=1, columnspan=2, sticky="ew")

tk.Label(input_excel_frame, text="Code column", anchor="w").grid(row=7, column=0, sticky="ew")
tk.Entry(input_excel_frame).grid(row=7, column=1, columnspan=2, sticky="ew")

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
ttk.Combobox(output_excel_frame, state="readonly", values=["Next empty column", "Next empty row"], textvariable=mode_cb_var).grid(row=3, column=1, columnspan=2, sticky="ew")

#tk.Label(output_excel_frame, text="Column Offset", anchor="w").grid(row=4, column=0, sticky="ew")
#tk.Entry(output_excel_frame, textvariable=column_offset_var).grid(row=4, column=1, columnspan=2, sticky="ew")

tk.Checkbutton(output_excel_frame, text="Include Name", variable=add_name_var).grid(row=5, column=0)

for widget in output_excel_frame.winfo_children():
    widget.grid_configure(padx=5, pady=3)

submit_buttons_frame = tk.Frame(file_info_frame)
submit_buttons_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=(0,5), sticky="news")

tk.Button(submit_buttons_frame, text="Delete from Table", command=delete_from_json).pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)
tk.Button(submit_buttons_frame, text="Update entry in Table", command=update_table).pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)
tk.Button(submit_buttons_frame, text="Save to Table", command=save_to_table).pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)

# Create the table to display saved entries
table_frame = tk.LabelFrame(frame)
#table_frame.grid(row=0, column=1, padx=5, pady=10, sticky="ns")
table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=tk.YES, padx=5, pady=5)

table = ttk.Treeview(table_frame, columns=(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11), show="headings", height="16")
#table.grid(row=1, column=0, padx=10, pady=10, sticky="news")
table.pack(side=tk.LEFT, fill=tk.X, expand=tk.YES)

table.heading(1, text="Name", anchor="center")
table.heading(2, text="Input FP", anchor="center")
table.heading(3, text="Input SN", anchor="center")
table.heading(4, text="Output FP", anchor="center")
table.heading(5, text="Output SN", anchor="center")
table.heading(6, text="Row Offset", anchor="center")
table.heading(7, text="Columns", anchor="center")
table.heading(8, text="Enabled", anchor="center")
table.heading(9, text="Mode", anchor="center")
table.heading(10, text="Add Name", anchor="center")
table.heading(11, text="Column Offset", anchor="center")

table.column("#1", anchor="w", stretch=True, width=75)
table.column("#2", anchor="w", stretch=True, width=80)
table.column("#3", anchor="w", stretch=True, width=75)
table.column("#4", anchor="w", stretch=True, width=80)
table.column("#5", anchor="w", stretch=True, width=75)
table.column("#6", anchor="w", stretch=True, width=75)
table.column("#7", anchor="w", stretch=True, width=75)
table.column("#8", anchor="w", stretch=True, width=75)
table.column("#9", anchor="w", stretch=True, width=100)
table.column("#10", anchor="w", stretch=True, width=75)
table.column("#11", anchor="w", stretch=True, width=75)

table.tag_configure(False, background="firebrick1")
table.tag_configure(True, background="lightgreen")

table_scroll_bar = tk.Scrollbar(table_frame, orient="vertical", command=table.yview)
table.configure(yscrollcommand=table_scroll_bar.set)
#table_scroll_bar.grid(row=1, column=1, sticky="ns")
table_scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)

table.bind('<<TreeviewSelect>>', on_table_select)

tk.Button(root, text="Save To File", command=save_to_file).pack(padx=5, pady=(0,5), side=tk.LEFT, fill=tk.X, expand=tk.YES)

tk.Button(root, text="Pull Data", command=lambda: column_extraction.run()).pack(padx=5, pady=(0,5), side=tk.LEFT, fill=tk.X, expand=tk.YES)

load_from_json()

# Run the application
root.mainloop()