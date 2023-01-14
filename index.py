from kiteTrade import *
import xlwings as xw
from constants import constants

def start():
    wb = xw.Book('data.xlsx')
    CREDENTIAL_SHEET = wb.sheets['CREDENTIAL']
    enctoken = CREDENTIAL_SHEET.range('B1').value
    if(enctoken):
        kite = KiteApp(enctoken=enctoken)
        LTP_SBIN = kite.ltp("NSE:SBIN")
        if(LTP_SBIN):
            print(constants.LOGIN_SUCCESS)
            CREDENTIAL_SHEET.range('B3').value = constants.LOGIN_SUCCESS
            CREDENTIAL_SHEET.range('B3').color = "#00FF00"
            CREDENTIAL_SHEET.range('B3').api.Font.ColorIndex = 1
            
        else:
            print(constants.INVALID_TOKEN)
            CREDENTIAL_SHEET.range('B3').value = constants.INVALID_TOKEN
            CREDENTIAL_SHEET.range('B3').color = "#FF0000"
            CREDENTIAL_SHEET.range('B3').api.Font.ColorIndex = 2
    else:
        print(constants.NO_TOKEN)
        CREDENTIAL_SHEET.range('B3').value = constants.NO_TOKEN
        CREDENTIAL_SHEET.range('B3').color = "#FF0000"
        CREDENTIAL_SHEET.range('B3').api.Font.ColorIndex = 2


start()
