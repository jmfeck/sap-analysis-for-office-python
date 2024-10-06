# -*- coding: utf-8 -*-
"""
@author: jfeck
@company: montblanc

"""

import sys
import os
import win32com.client
#import time

path_script = os.path.realpath(__file__)
#path_script = os.path.join(os.getcwd(), 'excel_update_application.py')
path_project = os.path.dirname(path_script)

file_mapping_btq_galery_merged = sys.argv[1]
path_mapping_gbu_per_user = os.path.join(path_project, file_mapping_btq_galery_merged)

# Start an instance of Excel
excel = win32com.client.DispatchEx("Excel.Application")

# Optional, e.g. if you want to debug
excel.WindowState = -4137
excel.Visible = True


wb = excel.Workbooks.Open(path_mapping_gbu_per_user)

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

