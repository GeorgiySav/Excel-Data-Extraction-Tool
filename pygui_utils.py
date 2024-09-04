import dearpygui.dearpygui as dpg
import xdialog

def create_file_input(input_tag, entry_label, button_label="Browse"):
    def button_callback():
        f = xdialog.open_file("Title here", filetypes=[("Excel Files", "*.xlsx")], multiple=False)
        dpg.set_value(input_tag, f)

    dpg.add_input_text(tag=input_tag, hint="Select a file", label=entry_label, readonly=True)
    dpg.add_button(label=button_label, callback=button_callback)


def add_error_msg_to_console(message):
    dpg.add_text(message, parent="console_window", color=(196, 43, 43))
    dpg.set_y_scroll("console_window", dpg.get_y_scroll_max("console_window"))


def add_valid_msg_to_console(message):
    dpg.add_text(message, parent="console_window", color=(69, 214, 69))
    dpg.set_y_scroll("console_window", dpg.get_y_scroll_max("console_window"))