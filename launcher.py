import os
import re
import sys
import logging
import threading
import traceback
import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import Checkbutton
from tkinter import Label

from docx.shared import Inches
from modify_docx import modify_docx
# CONSTANTS section

# if no XPT run value = '', else value = 'XPT-'. Used to generate docx report name and docx tags
# XPT_RUN = ''
# image width/height
IMAGE_WIDTH_VERTICAL = Inches(7.2)
IMAGE_WIDTH_HORIZONTAL = Inches(10.5)
#current directory
DIRECTORY = os.getcwd()
# filenames
FILENAME_DOCX_TEMPLATE = 'FTS_report_template.docx'
FILENAME_PPTX = ''

# look for ppt in the current directory
for file in os.listdir(DIRECTORY):
    if file.endswith(".pptx") and not file.startswith("~"):
        print(file)
        FILENAME_PPTX = file

# redirecting traceback during script execution to a file

# clear error_log file contents
open("error_log.txt", "w").close()
# Set up a logging handler to write to a file
log_file = "error_log.txt"
handler = logging.FileHandler(log_file, mode="a")
handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s %(levelname)s: %(message)s')
handler.setFormatter(formatter)
logging.getLogger().addHandler(handler)

# Redirect stdout and stderr to the logging module
sys.stdout = handler
sys.stderr = handler

# Lock to prevent duplicate log messages
log_lock = threading.Lock()

# Exception hook function to log uncaught exceptions
def log_uncaught_exception(exctype, value, tb):
    # Acquire the lock to prevent duplicate log messages
    log_lock.acquire()
    logging.error("Uncaught exception occurred:")
    logging.error(''.join(traceback.format_exception(exctype, value, tb)))
    log_lock.release()

# Install the exception hook
sys.excepthook = log_uncaught_exception

# Example script code that generates an error
# raise Exception("This is an error message")



# generate docx report
#modify_docx(FILENAME_PPTX, FILENAME_DOCX_TEMPLATE, FILENAME_DOCX_REPORT, IMAGE_WIDTH_VERTICAL, IMAGE_WIDTH_HORIZONTAL)

# user interface
window = tb.Window(themename="superhero")
window.geometry('400x280')
window.resizable(False, False)
window.title('FTS Wizard v1.0')
XPT_var = tk.StringVar()
MDT_var = tk.StringVar()
ORA_var = tk.StringVar()


app_label = tb.Label(text="FTS Wizard", font=("Helvetica", 20), bootstyle="success")
developer_label = tb.Label(text="Developed by Stanislav Kuzmin (SKuzmin2@slb.com)", font=("Helvetica", 8), bootstyle="default")

pretest_label = tb.Label(window, text="Pretests were done with:")

# used to generate text for <PRETEST_TOOL> placeholder in the report
XPT_checkbox = tb.Checkbutton(window, text = "XPT", variable =XPT_var, onvalue = "XPT-", offvalue = "", bootstyle="success")
MDT_checkbox = tb.Checkbutton(window, text = "MDT", variable =MDT_var, onvalue = "MDT-", offvalue = "", bootstyle="success")
ORA_checkbox = tb.Checkbutton(window, text = "ORA", variable =ORA_var, onvalue = "ORA-", offvalue = "", bootstyle="success")

progress_bar = tb.Progressbar(window, value=0, length=200, style='success.Striped.Horizontal.TProgressbar')
progress_lbl = tb.Label(window, text='')

btn_run_button = tb.Button(master=window,
                           text='Generate report',
                           command= lambda: modify_docx(
                               FILENAME_PPTX,
                               FILENAME_DOCX_TEMPLATE,
                               # FILENAME_DOCX_REPORT
                               re.search('.+(?=in_)', FILENAME_PPTX).group() + 'in_' + XPT_var.get() + \
                                    re.search('(?<=in_)(.+)(?=_Sampling)', FILENAME_PPTX).group() + '_Report',
                               IMAGE_WIDTH_VERTICAL,
                               IMAGE_WIDTH_HORIZONTAL,
                               XPT_var.get(),
                               [XPT_var.get(), MDT_var.get(), ORA_var.get()],
                               progress_bar,
                               progress_lbl
                           ),
                           bootstyle=(SUCCESS, OUTLINE)
                           )

app_label.place(x=200, y=30, anchor=CENTER)
pretest_label.place(x=200, y=70, anchor=CENTER)
XPT_checkbox.place(x=100, y=100, anchor=CENTER)
MDT_checkbox.place(x=200, y=100, anchor=CENTER)
ORA_checkbox.place(x=300, y=100, anchor=CENTER)
btn_run_button.place(x=200, y=150, anchor=CENTER)
progress_bar.place(x=200, y=200, anchor=CENTER)
progress_lbl.place(x=200, y=230, anchor=CENTER)
developer_label.place(x=200, y=260, anchor=CENTER)

window.mainloop()