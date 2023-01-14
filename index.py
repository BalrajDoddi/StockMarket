from kiteTrade import *

enctoken = "" # enctoken extract from Zerodha website
kite = KiteApp(enctoken=enctoken)

print(kite.margins())
print(kite.orders())
print(kite.positions())