from kiteTrade import *
import xlwings as xw
import pandas as pd
import time
from constants import constants

def updateInstruments(kite, wb):
    INSTRUMENTS = kite.instruments()
    INSTRUMENTS_SHEET = wb.sheets['INSTRUMENTS']
    INSTRUMENTS_SHEET.range('A1').value = pd.DataFrame(INSTRUMENTS)
    print(constants.INSTRUMENTS_UPDATED)


def startTradingProcess(kite, wb):
    while True:
        DATA_SHEET = wb.sheets('DATA')

        for i in range(2, 100):
            DATA = DATA_SHEET.range(f'A{i}:B{i}').value

            if(DATA[0] and DATA[1]):
                param = f'{DATA[0]}:{DATA[1]}'
                LTP = kite.ltp(param)

                if(LTP):
                    DATA_SHEET.range(f'C{i}').value = LTP[param]['instrument_token']
                    DATA_SHEET.range(f'D{i}').value = LTP[param]['last_price']
                    DATA_SHEET.range(f'K{i}').value = ''
                else:
                    print(constants.INCORRECT_SYMB)
                    DATA_SHEET.range(f'K{i}').value = constants.INCORRECT_SYMB
                    DATA_SHEET.range(f'K{i}').api.Font.ColorIndex = 3
            else:
                DATA_SHEET.range(f'C{i}').value = ''
                DATA_SHEET.range(f'D{i}').value = ''
        # time.sleep(1)

def updateHoldings(kite, wb):
    HOLDINGS = kite.holdings()
    HOLDINGS_SHEET = wb.sheets['HOLDINGS']

    HOLDINGS_SHEET.range('A1:A4').value = ['Symbol', 'exchange', 'instrument_token', 'average_price', 'last_price', 'pnl']

    for i in range(len(HOLDINGS)):
        # print(HOLDINGS[i])
        HOLDINGS_SHEET.range(f'A{i + 2}').value = HOLDINGS[i]['tradingsymbol']
        HOLDINGS_SHEET.range(f'B{i + 2}').value = HOLDINGS[i]['exchange']
        HOLDINGS_SHEET.range(f'C{i + 2}').value = HOLDINGS[i]['instrument_token']
        HOLDINGS_SHEET.range(f'D{i + 2}').value = HOLDINGS[i]['average_price']
        HOLDINGS_SHEET.range(f'E{i + 2}').value = HOLDINGS[i]['last_price']
        HOLDINGS_SHEET.range(f'F{i + 2}').value = HOLDINGS[i]['pnl']


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
            # updateHoldings(kite, wb)
            updateInstruments(kite, wb)
            startTradingProcess(kite, wb)
            
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
