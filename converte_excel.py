from fileinput import filename
import os
import win32com.client as client

excel = client.Dispatch("excel.application")
file = "Demonstrativo Gross_20220922.xlsb"

filename, fileextension = os.path.splitext(file)
wb = excel.Workbooks.Open(os.getcwd()  + "/" +  file)
output_path = os.getcwd() + "/" + filename
wb.SaveAs(output_path,51,ConflictResolution="xlLocalSessionChanges")
wb.Close()

excel.Quit()
