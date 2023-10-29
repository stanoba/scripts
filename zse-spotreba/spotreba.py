# pip3 install pandas
# pip3 install xlrd
# pip3 install tabulate

import pandas as pd
import math
from tabulate import tabulate, SEPARATING_LINE

##################################### Tariffs ####################################
# stav ku 31.8.2023
kwh = 0.08330000
monthly_pay= 1.50000000
spotr_dan = 0.00000000

sys_sluzby = 0.00629760
sys_prev = 0.01590000
mes_platba = 4.58070000
var_zlozka = 0.01300500
straty = 0.01146600
jadr_fond = 0.00327000

###################################### CODE ######################################
columns = ['date','time','hour_consumption']
df = pd.read_excel('export.xls', parse_dates=True, sheet_name='Nameraná hodinová práca', skiprows=10, usecols=[1,2,3], header=None, names = columns )


df2 = df.sum(numeric_only=True)
total_consumption = math.ceil(df2['hour_consumption'])
date_from = df.iloc[0, 0]
date_to = df.tail(1).iloc[0, 0]

days = date_to.split('.')
days = int(days[0])
ZSD = (kwh*total_consumption)+monthly_pay
ZSE = (sys_sluzby*total_consumption) + (sys_prev*total_consumption) + mes_platba + (var_zlozka*total_consumption) + (straty*total_consumption) + (jadr_fond*total_consumption)


df2 = pd.DataFrame([
[date_from, date_to, days,"Cena za elektrinu", total_consumption, "kWh", kwh, (kwh*total_consumption)],
[date_from, date_to, days,"Mesačná platba za jedno odberné miesto", "1.00", monthly_pay, "1 M",monthly_pay],
[date_from, date_to, days, "Spotrebná daň", total_consumption, "kWh", spotr_dan, (spotr_dan*total_consumption)],
SEPARATING_LINE,
["Spolu bez DPH","","","","","","",str(round(ZSD,2))+" €"]
],
columns=['Dátum od', 'Dátum do', 'Dni', 'Názov položky', 'Množstvo', 'Cena bez DPH\n (€/mer.j.)', 'Obdobie\npre cenu', 'Suma bez\nDPH (€)'])

df3 = pd.DataFrame([
[date_from, date_to, days,"Platba za systémové služby", total_consumption, "kWh", sys_sluzby, round((sys_sluzby*total_consumption),2)],
[date_from, date_to, days,"Platba za prevádzkovanie systému", total_consumption, "kWh", sys_prev, round((sys_prev*total_consumption),2)],
[date_from, date_to, days, "Pevná mesačná zložka tarify za 1 OM", "1 Mesiac", mes_platba, "1 M",round(mes_platba,2)],
[date_from, date_to, days, "Variabilná zložka tarify za distribúciu", total_consumption, "kWh", var_zlozka, round((var_zlozka*total_consumption),2)],
[date_from, date_to, days, "Platba za straty elektr.pri distr.el.", total_consumption, "kWh", straty, round((straty*total_consumption),2)],
[date_from, date_to, days, "Odvod do jadrového fondu", total_consumption, "kWh", jadr_fond, round((jadr_fond*total_consumption),2)],
SEPARATING_LINE,
["Spolu bez DPH","","","","","","",str(round(ZSE,2))+" €"]
],
columns=['Dátum od', 'Dátum do', 'Dni', 'Názov položky', 'Množstvo', 'Cena bez DPH\n (€/mer.j.)', 'Obdobie\npre cenu', 'Suma bez\nDPH (€)'])

##########################################################################################################################################
print()
print("Dodávka elektriny: produkt DomovKlasik - DD2 (fakturačné položky)")

print(tabulate(df2, headers = 'keys', tablefmt = 'psql', showindex="False"))
#print(df2)

print()
print("Distribúcia a regulované poplatky: produkt D2 (fakturačné položky)")
#print(df3)
print(tabulate(df3, headers = 'keys', tablefmt = 'psql', showindex="False"))

print("Spolu celkovo bez DPH",round(ZSD + ZSE,2),"€")
print("Spolu celkovo s DPH",round((ZSD + ZSE)*1.2,2),"€")
print()
