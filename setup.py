import cx_Freeze
import sys
import os 
base = None

if sys.platform == 'win32':
    base = "Win32GUI"

os.environ['TCL_LIBRARY'] = r"C:\Users\Star X-Prt\AppData\Local\Programs\Python\Python311\tcl\tcl8.6"
os.environ['TK_LIBRARY'] = r"C:\Users\Star X-Prt\AppData\Local\Programs\Python\Python311\tcl\tk8.6"

executables = [cx_Freeze.Executable("St Details.py", base=base, icon="icon.ico")]


cx_Freeze.setup(
    name = "St Details",
    options = {"build_exe": {"packages":["tkinter","os"], "include_files":["icon.ico","Student Images",'tcl86t.dll','tk86t.dll', 'imges','Student_data.xlsx']}},
    version = "1.0",
    description = "Student Managment System | Developed By Sonu Yadav",
    executables = executables
    )
