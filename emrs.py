import pandas as pd
import jdatetime
import re

def to_jalali(x):
    if isinstance(x, str):
        m = re.search(r"\b(14\d\d)[-/](0[1-9]|1[0-2])[-/](0[1-9]|[12]\d|3[01])\b", x)
        if m:
            text = x.split('/')
            text =[int(x) for x in text]
            return jdatetime.date(year=text[0],month=text[1],day=text[2])
    return None
#print(to_jalali('1403/01/02'))

emrs_parseh = pd.read_excel('شاخص EMRS (سورت شده).xlsx',header=None)
emrs_parseh[0] = pd.to_numeric(emrs_parseh[0], errors='coerce')
emrs_parseh = emrs_parseh[emrs_parseh[0].notna()]
emrs_parseh.reset_index(inplace=True)
#print(emrs_parseh.tail(20))

dastorkar = pd.read_excel('تاریخچه دستور کارها.xlsx')

dastorkar['تاریخ شروع'] = dastorkar['تاریخ شروع'].apply(to_jalali)
dastorkar['تاریخ اتمام'] = dastorkar['تاریخ اتمام'].apply(to_jalali)
delta = dastorkar['تاریخ اتمام'] - dastorkar['تاریخ شروع']

print([de.days for de in delta if isinstance(de, jdatetime.datetime)])


