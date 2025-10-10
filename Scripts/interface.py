# import dearpygui.dearpygui as dpg
# import dearpygui.demo as demo
#
# dpg.create_context()
# dpg.create_viewport(title='Custom Title', width=600, height=600)
#
# demo.show_demo()
#
# dpg.setup_dearpygui()
# dpg.show_viewport()
# dpg.start_dearpygui()
# dpg.destroy_context()

# import dearpygui.dearpygui as dpg
#
# dpg.create_context()

# def callback(sender, app_data):
#     print('OK was clicked.')
#     print("Sender: ", sender)
#     print("App Data: ", app_data)
#
# def cancel_callback(sender, app_data):
#     print('Cancel was clicked.')
#     print("Sender: ", sender)
#     print("App Data: ", app_data)
#
# dpg.add_file_dialog(
#     directory_selector=True, show=False, callback=callback, tag="file_dialog_id",
#     cancel_callback=cancel_callback, width=700 ,height=400)
#
# with dpg.window(label="Tutorial", width=800, height=300):
#     with dpg.file_dialog(label="Demo File Dialog", width=600, height=400, show=False, callback=lambda s, a, u : print(a["selections"].values()), tag="__demo_filedialog"):
#         dpg.add_file_extension(".*", color=(255, 255, 255, 255))
#         dpg.add_file_extension("Source files (*.cpp *.h *.hpp){.cpp,.h,.hpp}", color=(0, 255, 255, 255))
#         dpg.add_file_extension(".xlsx", color=(0, 255, 0, 255), custom_text="EXCEL")
#         dpg.add_file_extension(".pdf", color=(255, 0, 0, 255), custom_text="PDF")
#         dpg.add_file_extension(".py", color=(0, 255, 255, 255))
#         #dpg.add_button(label="Button on file dialog")
#
#     dpg.add_button(label="Show File Selector", user_data=dpg.last_container(), callback=lambda s, a, u: dpg.configure_item(u, show=True))
#
# dpg.create_viewport(title='Custom Title', width=800, height=600)
# dpg.setup_dearpygui()
# dpg.show_viewport()
# dpg.start_dearpygui()
# dpg.destroy_context()
#
# import dearpygui.dearpygui as dpg
#
# dpg.create_context()
#
# with dpg.window(label="Tutorial", pos=(20, 50), width=275, height=225) as win1:
#     t1 = dpg.add_input_text(default_value="some text")
#     t2 = dpg.add_input_text(default_value="some text")
#     with dpg.child_window(height=100):
#         t3 = dpg.add_input_text(default_value="some text")
#         dpg.add_input_int()
#     dpg.add_input_text(default_value="some text")
#
# with dpg.window(label="Tutorial", pos=(320, 50), width=275, height=225) as win2:
#     dpg.add_input_text(default_value="some text")
#     dpg.add_input_int()
#
# with dpg.theme() as global_theme:
#
#     with dpg.theme_component(dpg.mvAll):
#         dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (255, 140, 23), category=dpg.mvThemeCat_Core)
#         dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 5, category=dpg.mvThemeCat_Core)
#
#     with dpg.theme_component(dpg.mvInputInt):
#         dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (140, 255, 23), category=dpg.mvThemeCat_Core)
#         dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 5, category=dpg.mvThemeCat_Core)
#
# dpg.bind_theme(global_theme)
#
# dpg.show_style_editor()
#
# dpg.create_viewport(title='Custom Title', width=800, height=600)
# dpg.setup_dearpygui()
# dpg.show_viewport()
# dpg.start_dearpygui()
# dpg.destroy_context()
import pdfplumber as pl
import re
from collections import namedtuple
import xlrd
import pandas as pd
import openpyxl
from copy import copy
from icecream import ic
import dearpygui.dearpygui as dpg



create_name_table("","signals.xlsx", enable_export=True)