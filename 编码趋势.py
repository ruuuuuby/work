# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 13:53:42 2021

@author: 50256279
"""

import pandas as pd
import numpy as np
import matplotlib.mlab as mlab
import matplotlib.pyplot as plt
from statsmodels.graphics.tsaplots import plot_acf,plot_pacf
import statsmodels.api as sm
import warnings
warnings.filterwarnings("ignore")

path='C:/Users/50256279/Desktop/4.13'
file=pd.read_excel(path+'/16输注泵旧工艺代号产品编码.XLSX',dtype=str)
a=file['有效起始期'].str.split('-')
file['YEAR']=[i[0] for i in a]
file['MONTH']=[i[1] for i in a]
file['DATE']=file['YEAR']+file['MONTH']

b=file.groupby(['DATE'])['物料编码'].count()
b=b.reset_index()


#2018后情况
plt.figure(figsize=(20,12),dpi=500)
plt.bar(b[b['DATE']>'201712']['DATE'],b[b['DATE']>'201712']['物料编码'],0.4,color="green")
plt.xticks(rotation=270)
plt.grid(True)



#总体情况
plt.figure(figsize=(20,12),dpi=500)
plt.xticks(rotation=360)
plt.bar(b['DATE'],b['物料编码'],0.4,color="green")
plt.xticks(rotation=270)
plt.grid(True)

plot_acf(b['物料编码'])
plot_pacf(b['物料编码'])

mod=sm.tsa.statespace.SARIMAX(b[b['DATE']>'201712']['物料编码'],order=(1,0,2),seasonal_order=(1,0,1,12))
results=mod.fit()
c=results.predict(start=102,end=109)
c=c.reset_index()
c.drop(columns='index',inplace=True)
b['FORECAST']=c
k=['202105','202106','202107','202108','202109','202110','202111','202112']

new2=pd.DataFrame({'DATE':k})
new2['物料编码']=c

new2.to_excel(path+'/1.xlsx')