from openpyxl import load_workbook
from openpyxl import Workbook
import re

sen = 'Hi, my name is Jeter.'

print(re.search('jeter', sen) == None)
