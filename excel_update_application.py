import sys
import os
import win32com.client

script_path = os.path.realpath(__file__)
#path_script = os.path.join(os.getcwd(), 'excel_update_application.py')
project_path = os.path.dirname(script_path)

# File to be updated 
file_name = sys.argv[1]
file_path = os.path.join(project_path, file_name)

# Start Excel Instance
excel = win32com.client.DispatchEx("Excel.Application")

# Command to open excel
# I dont have a work arround to make it work being false
excel.Visible = True

wb = excel.Workbooks.Open(file_path)

analysis_addin = excel.COMAddIns("SapExcelAddIn")
analysis_addin.Connect = False
analysis_addin.Connect = True

excel.Application.Run("SAPExecuteCommand", "Refresh")

# Save and close excel
wb.Save()
wb.Close()

# Quit
excel.Quit()
del excel

# Quit
excel.Quit()
del excel

