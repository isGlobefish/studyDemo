import pandas as pd
import numpy as np

data = pd.read_excel('C:/Users/Zeus/Desktop/无法查看单价的商品明细清单.xlsx')

data['销售日期'] = data['销售日期'].astype('datetime64')
data.to_excel('C:/Users/Zeus/Desktop/save0.xlsx')
data.dtypes


for i in range(len(data)):
    print('"000000000' + str(data.loc[i,"商品编码"]) + '"')