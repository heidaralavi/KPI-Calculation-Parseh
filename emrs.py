import pandas as pd
import jdatetime
import re

# function for jalalli datetime
def to_jalali(x):
    if isinstance(x, str):
        m = re.search(r"\b(14\d\d)[-/](0[1-9]|1[0-2])[-/](0[1-9]|[12]\d|3[01])\b", x)
        if m:
            text = x.split('/')
            text =[int(x) for x in text]
            return jdatetime.date(year=text[0],month=text[1],day=text[2])
    return None
#print(to_jalali('1403/01/02'))

# reading parseh output results
emrs_parseh = pd.read_excel('شاخص EMRS (سورت شده).xlsx',header=None)

# filltering empty columns and cells
emrs_parseh[0] = pd.to_numeric(emrs_parseh[0], errors='coerce')*100
emrs_parseh = emrs_parseh[emrs_parseh[0].notna()]
emrs_parseh.drop(labels=[2,3,5,7,8,9,10,12,13],inplace=True,axis=1)
emrs_parseh.reset_index(inplace=True,drop=True)
columns_name = {0:'EMRS',1:'تعداد دستورکارهای EM',4:'تعداد EMهای اتمام یافته کمتراز ۳روز',6:'AppTag تجهیز',11:'نام تجهیز',14:'ردیف'}
emrs_parseh.rename(columns=columns_name,inplace=True)
#saving parseh output to excell
emrs_parseh.to_excel('parseh.xlsx',index=False)
#print(emrs_parseh.tail(20))

#reading all workorder from parseh
dastorkar = pd.read_excel('تاریخچه دستور کارها.xlsx')
#print(dastorkar.shape)

#filltering data
dastorkar['تاریخ شروع'] = dastorkar['تاریخ شروع'].apply(to_jalali)
dastorkar['تاریخ اتمام'] = dastorkar['تاریخ اتمام'].apply(to_jalali)
from_date = jdatetime.date(year=1403,month=5,day=1)
to_date = jdatetime.date(year=1403,month=11,day=1)

mask_1 = dastorkar['وضعیت'].eq('اتمام یافته')
#print(mask_1)
dastorkar = dastorkar[mask_1]
mask_2 = dastorkar['نوع فعالیت'].eq('اضطراری')
#print(mask_2)
dastorkar = dastorkar[mask_2]
mask_3 = dastorkar['تاریخ شروع'].between(from_date,to_date)
dastorkar = dastorkar[mask_3]
#print(dastorkar.head())

#calculate task durations
delta = dastorkar['تاریخ اتمام'] - dastorkar['تاریخ شروع']
delta_days = []
smaler_than_3 = []
for item in delta:
    #print(type(item))
    if isinstance(item, jdatetime.timedelta):
        if item.days <=3 :
            smaler_than_3.append(1)
        else:
            smaler_than_3.append(0)
        delta_days.append(item.days)
        #print(item.days)
    else:
        delta_days.append(None)
        smaler_than_3.append(None)

#print(len(delta_days))

#add to dataframe
dastorkar['delta'] = delta_days
dastorkar['smaler_than_3'] = smaler_than_3

#calculating
temp1 = dastorkar.groupby('AppTAG')['smaler_than_3'].sum()
temp2 = dastorkar.groupby('AppTAG')['شماره دستور کار'].nunique()
temp = pd.merge(temp1,temp2,on='AppTAG')
temp['emrs'] = round(100.0*temp['smaler_than_3']/temp['شماره دستور کار'],2)
temp.sort_values(by='emrs',ascending=True,inplace=True)
#print(temp.head())

#rename columns
columns_newname = {'شماره دستور کار':'تعداد دستورکارهای اضطراری'}
temp.rename(columns=columns_newname,inplace=True)
#saving
temp.to_excel('check_point.xlsx')

#dastorkar.to_excel('temp_t.xlsx')

