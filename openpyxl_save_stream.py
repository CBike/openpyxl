from tempfile import NamedTemporaryFile
from openpyxl import Workbook


wb = load_workbook()
wb.template = True

wb.save('document_template.xltx')

