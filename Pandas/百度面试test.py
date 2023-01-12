'''
逝者如斯夫, 不舍昼夜 -- 孔夫子
@Auhor    : Dohoo Zou
Project   : gitCode
FileName  : 百度面试test.py
IDE       : PyCharm
CreateTime: 2023-01-12 10:51:36
'''
import pandas as pd

data = {
    '订单号': ['19160394', '&nbsp19140861', '&nbsp19160333'],
    '物流费用': [0.2513, 0.1513, 0.1513]
}
df1 = pd.DataFrame(data)

fpath = '/Users/dohozou/Desktop/Code/gitCode/Pandas/dataFiles/'
df1 = pd.read_excel(fpath + 'df1.xlsx', sheet_name=0)
df2 = pd.read_excel(fpath + 'df2.xlsx', sheet_name=0)
df3 = pd.read_excel(fpath + 'df3.xlsx', sheet_name=0)
df4 = pd.read_excel(fpath + 'df4.xlsx', sheet_name=0)
df5 = pd.read_excel(fpath + 'df5.xlsx', sheet_name=0)
df6 = pd.read_excel(fpath + 'df6.xlsx', sheet_name=0)
df7 = pd.read_excel(fpath + 'df7.xlsx', sheet_name=0)
df8 = pd.read_excel(fpath + 'df8.xlsx', sheet_name=0)

# 题目1
df1.loc[:, '订单号'] = df1['订单号'].str.replace('&nbsp', '').astype('int32')

# 题目2
df3 = pd.merge(df1, df2, on='订单号', how='left')

# 题目3
df4.loc[df4['订单号'].str.endswith('-1'), '采购成本'] = 0

# 题目4
df5.loc[:, '负责人（中）'] = df5.apply(lambda x: '杨天天' if x['负责人（拼）'] in ['ytt', 'yttt'] else '李月', axis=1)

# 题目5
df6.groupby(['负责人']).sum()

# 题目6
purchase_df = df7.loc[~df7['采购单号'].isin(['出库单', '库存调整']), :]

# 题目7
df8['员工代码'] = df8['商品名称'].str.extract(r'(g\d{1}[a-zA-Z]{2,3})', expand=False).str.strip()
