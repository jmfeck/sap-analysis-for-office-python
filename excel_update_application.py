# -*- coding: utf-8 -*-
"""
@author: jfeck

"""

import sys
import os
import win32com.client
#import time

path_script = os.path.realpath(__file__)
#path_script = os.path.join(os.getcwd(), 'excel_update_application.py')
path_project = os.path.dirname(path_script)

# File to be updated 
file_path = sys.argv[1]

# Start Excel Instance
excel = win32com.client.DispatchEx("Excel.Application")

# Optional, e.g. if you want to debug
excel.Visible = True

wb = excel.Workbooks.Open(file_path)

analysis_addin = excel.COMAddIns("SapExcelAddIn")
analysis_addin.Connect = False
analysis_addin.Connect = True

excel.Application.Run("SAPExecuteCommand", "Refresh")

# Refresh all data connections.
#wb.RefreshAll()
#time.sleep(5)

# Save and close excel
wb.Save()
wb.Close()

# Quit
excel.Quit()
del excel

