import xlwings as xw
import pandas as pd

app = xw.App(visible=False)
wb0 = xw.Book('상품목록 (2).xlsx')

pl = wb0.sheets[0]
pl.range('B:D').copy()
pl.range('AQ1').paste()
pl.range('AM:AP').delete()
pl.range('AE:AJ').delete()
pl.range('R:AA').delete()
pl.range('G:H').delete()
pl.range('A:D').delete()


wb = xw.Book('상품팀일일매출_05.xlsx')

sht1 = wb.sheets[10]
sht1.range('B:AI').clear()
sht1 = wb.sheets[11]
sht1.range('B:X').clear()
sht1 = wb.sheets[12]
sht1.range('A:K').clear()
sht1 = wb.sheets[13]
sht1.range('A:K').clear()
sht1 = wb.sheets[14]
sht1.range('A:K').clear()
sht1 = wb.sheets[15]
sht1.range('A:S').clear()
sht1.range('A1').paste()

pl.range('A:S').copy()
pl.
