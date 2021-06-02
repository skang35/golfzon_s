import xlwings as xw
import pandas as pd
from datetime import date, timedelta

app = xw.App(visible=False, add_book=False)
wb = xw.Book('C:\\Users\owner\PycharmProjects\pythonProject\dailySales\상품팀일일매출_05.xlsx')
# 입반출(사입)
sht1 = wb.sheets[10]
sht1.range('B:AI').clear()
# 통합수불
sht1 = wb.sheets[11]
sht1.range('B:X').clear()
# BEST판매
sht1 = wb.sheets[12]
sht1.range('A:K').clear()
# 전일판매
sht1 = wb.sheets[13]
sht1.range('A:K').clear()
# 전월판매
sht1 = wb.sheets[14]
sht1.range('A:K').clear()
# 상품코드
sht1 = wb.sheets[15]
sht1.range('A:S').clear()

# 상품코드RAWDATA
# app = xw.App(visible=False)
wb0 = xw.Book('상품목록 (2).xlsx')
print(1)

pl = wb0.sheets[0]
print(2)
pl.range('B:D').copy()
pl.range('AQ:AS').paste()
pl.range('AM:AP').delete()
pl.range('AE:AJ').delete()
pl.range('R:AA').delete()
pl.range('G:H').delete()
pl.range('A:D').delete()
print(3)
pl.range('A:S').copy(sht1.range('A:A'))
print(4)
wb0.close()
print(5)
# 입반출(사입) RAW
wb0 = xw.Book('입고반출현황2021-06-01.xlsx')
pl = wb0.sheets[0]
wb0.sheets[0].range('A:AH').copy(wb.sheets[10].range('B:B'))
wb0.close()
# 통합수불 RAW
wb0 = xw.Book('통합수불현황_2021-06-01.xlsx')
pl = wb0.sheets[0]
wb0.sheets[0].range('A:W').copy(wb.sheets[11].range('B:B'))
wb0.close()
# BEST판매 RAW
wb0 = xw.Book('Anl_SalesStockClsCondition (1).xlsx')
pl = wb0.sheets[0]
wb0.sheets[0].range('A:K').copy(wb.sheets[12].range('A:A'))
wb0.close()
# 전일판매 RAW
wb0 = xw.Book('Anl_SalesStockClsCondition.xlsx')
pl = wb0.sheets[0]
wb0.sheets[0].range('A:K').copy(wb.sheets[13].range('A:A'))
wb0.close()
# 전월판매 RAW
wb0 = xw.Book('Anl_SalesStockClsCondition (2).xlsx')
pl = wb0.sheets[0]
wb0.sheets[0].range('A:K').copy(wb.sheets[14].range('A:A'))
wb0.close()
