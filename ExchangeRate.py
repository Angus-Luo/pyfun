# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas as pd
# %%
df = pd.read_excel('\\\\tpfile01\\8200_Accounting\\PowerBI\\Mapping_General\\Exchange.xlsx', encoding='utf-8', header=0)
# %%
def add_currency(dataframe,new_cur: str):
    #added two currency
    #added 3 currency
    temp_df = dataframe[dataframe['Currency'] == new_cur]
    temp_df.drop('Currency', axis=1)
    temp_df.rename(columns={'ExchangeToTWD':f'{new_cur}toTWD'}, inplace=True)
    return temp_df

# %%
usd = add_currency(df,'USD')
gbp = add_currency(df,'GBP')
rmb = add_currency(df,'RMB')
hkd = add_currency(df,'HKD')
eur = add_currency(df,'EUR')
# %%
for idf in [usd,gbp,rmb,hkd,eur]:
    df = pd.merge(df, idf.drop('Currency', axis=1) , how='left', left_on=['FS_Type', 'Year', 'MonthNo'], 
         right_on=['FS_Type', 'Year', 'MonthNo'])

# %%
# df = pd.merge(df, usd.drop('Currency', axis=1) , how='left', left_on=['FS_Type', 'Year', 'MonthNo'], 
#          right_on=['FS_Type', 'Year', 'MonthNo'])
# df = pd.merge(df, gbp.drop('Currency', axis=1) , how='left', left_on=['FS_Type', 'Year', 'MonthNo'], 
#          right_on=['FS_Type', 'Year', 'MonthNo'])
# df = pd.merge(df, hkd.drop('Currency', axis=1) , how='left', left_on=['FS_Type', 'Year', 'MonthNo'], 
#          right_on=['FS_Type', 'Year', 'MonthNo'])
# df = pd.merge(df, rmb.drop('Currency', axis=1) , how='left', left_on=['FS_Type', 'Year', 'MonthNo'], 
#          right_on=['FS_Type', 'Year', 'MonthNo'])
# %%
for cur in ['USD','GBP','HKD','RMB','EUR']:
    df[f'ExchangeTo{cur}'] = df['ExchangeToTWD'] / df[f'{cur}toTWD']
    df.drop([f'{cur}toTWD'], axis=1, inplace=True)

# %%
# df['ExchangeToUSD'] = df['ExchangeToTWD'] / df['USDtoTWD']
# df['ExchangeToGBP'] = df['ExchangeToTWD'] / df['GBPtoTWD']
# df['ExchangeToHKD'] = df['ExchangeToTWD'] / df['HKDtoTWD']
# df['ExchangeToRMB'] = df['ExchangeToTWD'] / df['RMBtoTWD']


# %%
# df.drop(['USDtoTWD', 'GBPtoTWD', 'HKDtoTWD', 'RMBtoTWD'], axis=1, inplace=True)

# %%
df['FS_Type2'] = '(' + df['FS_Type'] + ')'


# %%
df['Currency_Key'] = df['FS_Type'] + df['Year'].astype(str) + df['MonthNo'].astype(str) + df['Currency']


# %%
# Standard exchange rate format
df.to_csv('\\\\tpfile01\\8200_Accounting\\PowerBI\\Mapping_General\\ExchangeRate.csv', encoding='utf-8', index=False)
# %%
oldex = df.copy()


# %%
Orgmap = {'TWD':'TW', 'RMB':'CHINA', 'GBP':'EMEA', 'USD':'USA', 'HKD':'HK'}
Orgmap2 = {'TWD':'AP', 'RMB':'CHINA', 'GBP':'EMEA', 'USD':'USA', 'HKD':'HK'}
fsmap = {'A':'Actual', 'Q2F':'Q2 FCST', 'B':'Q0Budget', 'Q1F': 'Q1 FCST', 'Q3F': 'Q3 FCST', 
         'R': 'Rolling', 'Q3F_9A':'Q3F_9A', 'Q2F_6A':'Q2F_6A', 'Q3F_10A':'Q3F_10A', 'Q3F_11A':'Q3F_11A'}


# %%
oldex['FS Type2'] = '(' + oldex['FS_Type'] +')'
oldex['ExchangeRate'] = oldex['ExchangeToTWD']
oldex['Currency_Key'] = oldex['Currency'] + oldex['Year'].astype(str) + oldex['FS Type2'] + oldex['MonthNo'].astype(str)
oldex['Organization'] = oldex['Currency'].map(Orgmap)
oldex['Organization2'] = oldex['Currency'].map(Orgmap2)
oldex['FS Type3'] = oldex['FS_Type'].map(fsmap)
oldex['Currency_Key2'] = oldex['Organization2'] + oldex['Year'].astype(str) + oldex['MonthNo'].astype(str) + oldex['FS Type3'] 


# %%
ex = oldex[['Year', 'MonthNo', 'FS Type2','ExchangeToUSD',
       'ExchangeToGBP', 'ExchangeToRMB', 'ExchangeToHKD']][oldex['Currency'] == 'TWD']


# %%
oldex = pd.merge(oldex.drop(['ExchangeToUSD', 'ExchangeToGBP', 'ExchangeToRMB', 'ExchangeToHKD'],1) ,ex,how='left', on=['Year', 'MonthNo', 'FS Type2'])


# %%
oldex.rename(columns={'ExchangeToTWD':'ExchangeTWD', 'ExchangeToUSD' : 'ExchangeUSD' , 'ExchangeToGBP':'ExchangeGBP', 
                   'ExchangeToRMB':'ExchangeRMB','ExchangeToHKD':'ExchangeHKD'}, inplace=True)


# %%
oldex.drop(['Organization2', 'FS_Type','FS_Type2'], axis=1, inplace=True)


# %%
oldex.to_excel('\\\\tpfile01\\8200_Accounting\\Erin\\主管交辦事項\\Model DB\\Mapping\\ExchangeRate.xlsx', encoding='utf-8',
           index=False)


# %%
# Done

