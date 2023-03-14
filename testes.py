#pip install pillow
import os
import win32com.client as client
from PIL import ImageGrab

# make sure you use either a raw string OR escape characters for the backslash. 
# You CANNOT use forward slash format with `win32com`.
workbook_path = r'S:\Resultados\01_Relatorio Diario\1 - Base Eventos\02 - TENDÊNCIA\TEND_FIBRA_e_UPDATEs.xlsx'

# start an instance of Excel
excel = client.Dispatch('Excel.Application')

# open the workbook
wb = excel.Workbooks.Open(workbook_path)

# select the sheet you want... use whatever index method you want to use....

### by item
#sheet = wb.Sheets.Item(1)

### by index
#sheet = wb.Sheets[0]

### by name
sheet = wb.Sheets['UPDATE_TENDENCIA_VLL']

# copy the target range
copyrange= sheet.Range('A23:D45')
copyrange.CopyPicture(Appearance=1, Format=2)

# grab the saved image from the clipboard and save to working directory
ImageGrab.grabclipboard().save('paste.png')


# get the path of the current working directory and create image path
image_path = os.getcwd() + '\\paste.png'

# create a html body template and set the **src** property with `{}` so that we can use
# python string formatting to insert a variable
html_body = """
    <div>
          Please review the following report and response with your feedback.
    </div>
    <div>
        <img src={}></img>
    </div>
"""

# startup and instance of outlook
outlook = client.Dispatch('Outlook.Application')

# create a message
message = outlook.CreateItem(0)

# set the message properties
message.To = 'luiz.lobao@oi.net.br'
message.Subject = 'Please review!'
message.HTMLBody = html_body.format(image_path)

# display the message to review
message.Display()

# save or send the message
message.Send()