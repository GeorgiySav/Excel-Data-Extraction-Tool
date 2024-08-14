import dearpygui.dearpygui as dpg
from openpyxl.utils.cell import column_index_from_string

import columns_gui
import rows_gui
import parse_gui

dpg.create_context()

with dpg.window(tag="Primary Window"):
    with dpg.tab_bar():
        with dpg.tab(label="Columns"):
            columns_gui.columns_gui()    
        with dpg.tab(label="Rows"):
            rows_gui.rows_gui()
        with dpg.tab(label="Parse"):
            parse_gui.parse_gui()

    with dpg.child_window(menubar=True, horizontal_scrollbar=True, autosize_x=True, height=300, tag="console_window"):
        with dpg.menu_bar():
            dpg.add_menu(label="Status Reports")
    

dpg.create_viewport(title='Custom Title', width=900, height=900)
dpg.setup_dearpygui()

dpg.show_viewport()
dpg.set_primary_window("Primary Window", True)
dpg.start_dearpygui()

dpg.destroy_context()