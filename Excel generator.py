"""picking a spreadshet based on the month
changing a few cells and updating it
then saving with proper name
"""

from datetime import datetime
from datetime import timedelta
import win32com.client

trinta = [4, 6, 9, 11]
trinta_um = [1, 3, 5, 7, 8, 10, 12]

if datetime.now().month in trinta:
    MODELO = 'path/MODELO 30.xlsx'
elif datetime.now().month in trinta_um:
    MODELO = 'path/MODELO 31.xlsx'
else:
    MODELO = 'path/MODELO 29.xlsx'
try:
    EXCEL = win32com.client.Dispatch("EXCEL.Application")
    EXCEL.Visible = False
    PLANILHA = EXCEL.Workbooks.Open(MODELO)
    SHEET = PLANILHA.WorkSHEETs('sheet name')
    SHEET.Cells(2, 2).Value = datetime.now()
    SHEET.Cells(1, 2).Value = (datetime.now()+timedelta(days=3))
    ano = datetime.now().strftime('%y')
    mes = datetime.now().strftime('%m')
    SHEET.Cells(1, 5).Value = f'{ano}{mes} Planejamento Oficial'
    PLANILHA.RefreshAll()
    EXCEL.CalculateUntilAsyncQueriesDone()
    dia_aprest = (datetime.now()+timedelta(days=3)).strftime('%d')
    mes_aprest = (datetime.now()+timedelta(days=3)).strftime('%m')
    ano_aprest = (datetime.now()+timedelta(days=3)).strftime('%y')
    PLANILHA.SaveAs(Filename=f"path\\Acompanhamento Semanal de Movimentações_{
                    dia_aprest}-{mes_aprest}-{ano_aprest}.xlsx")
    PLANILHA.Close()
    EXCEL.Quit()
except Exception as e:
    print(e)
finally:
    SHEET = None
    PLANILHA = None
    EXCEL = None
