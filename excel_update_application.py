import sys
import os
import win32com.client

def main():
    try:
        # Determine script and project paths
        script_path = os.path.realpath(__file__)
        project_path = os.path.dirname(script_path)

        # File to be updated
        if len(sys.argv) < 2:
            print("Usage: python excel_update_application.py <file_name>")
            sys.exit(1)
        
        input_excel_filename = sys.argv[1]
        file_path = os.path.join(project_path, input_excel_filename)

        if not os.path.exists(file_path):
            print(f"Error: File '{file_path}' not found.")
            sys.exit(1)

        # Start Excel instance
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True

        # Open the workbook
        workbook = excel.Workbooks.Open(file_path)

        # Refresh the SAP add-in
        analysis_addin = excel.COMAddIns("SapExcelAddIn")
        analysis_addin.Connect = False
        analysis_addin.Connect = True

        excel.Application.Run("SAPExecuteCommand", "Refresh")

        # Save and close the workbook
        workbook.Save()
        workbook.Close()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit Excel application
        if 'excel' in locals():
            excel.Quit()
            del excel

if __name__ == "__main__":
    main()
