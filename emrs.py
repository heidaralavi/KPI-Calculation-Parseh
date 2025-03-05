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

emrs_parseh.drop(labels=[2,3,5,7,8,9,10,12,13],inplace=True,axis=1)
emrs_parseh.reset_index(inplace=True,drop=True)
emrs_parseh.to_excel('parseh.xlsx',index=False)
#print(emrs_parseh.tail(20))

dastorkar = pd.read_excel('تاریخچه دستور کارها.xlsx')
print(dastorkar.shape)
dastorkar['تاریخ شروع'] = dastorkar['تاریخ شروع'].apply(to_jalali)
dastorkar['تاریخ اتمام'] = dastorkar['تاریخ اتمام'].apply(to_jalali)
from_date = jdatetime.date(year=1403,month=5,day=1)
to_date = jdatetime.date(year=1403,month=11,day=1)
mask_1 = dastorkar['تاریخ شروع'].between(from_date,to_date)
dastorkar = dastorkar[mask_1]
mask_2 = dastorkar['نوع فعالیت'].eq('اضطراری')
print(mask_2)
dastorkar = dastorkar[mask_2]
#print(dastorkar.head())
delta = dastorkar['تاریخ اتمام'] - dastorkar['تاریخ شروع']
delta_days = []
log3 = []
for item in delta:
    #print(type(item))
    if isinstance(item, jdatetime.timedelta):
        if item.days <=3 :
            log3.append(1)
        else:
            log3.append(0)
        delta_days.append(item.days)
        #print(item.days)
    else:
        delta_days.append(None)
        log3.append(None)

#print(len(delta_days))
dastorkar['delta'] = delta_days
dastorkar['LOG3'] = log3


test1 = dastorkar.groupby('AppTAG')['LOG3'].sum()
test2 = dastorkar.groupby('AppTAG')['شماره دستور کار'].nunique()
test3 = pd.merge(test1,test2,on='AppTAG')
test3['emrs'] = round(100.0*test3['LOG3']/test3['شماره دستور کار'],2)
test3.sort_values(by='emrs',ascending=True,inplace=True)
print(test3.head())
test3.to_excel('check_point.xlsx')
#dastorkar.to_excel('temp_t.xlsx')

