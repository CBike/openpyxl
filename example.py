from openpyxl import Workbook


wb = Workbook()  # 새 워크북 생성

ws = wb.active  # 현재 활성화 된 sheet 가져옴 , ws = wb['Sheet']

ws.title = 'NadoSheet'
wb.save('nadoCoding_sample.xlsx')
wb.close()