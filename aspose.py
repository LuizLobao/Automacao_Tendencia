import  jpype     
import  asposecells     

jpype.startJVM() 
from asposecells.api import Workbook

workbook = Workbook("Demonstrativo Gross_20220922.xlsb")
workbook.save("Demonstrativo Gross_20220922.xlsx")
jpype.shutdownJVM()