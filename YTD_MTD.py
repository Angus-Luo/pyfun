# %%
import pandas as pd
import datetime
# %%
df = pd.read_csv(r'\\tpfile01\8200_Accounting\PowerBI\Mapping_General\calendar.csv')

# %%
today = datetime.datetime.today()
# %%
datetime.datetime.today()
# %%

year_begin = datetime.datetime(today.year, 1, 1)
month_begin = datetime.datetime(today.year,today.month,1)
# %%
df['Date'] = pd.to_datetime(df['Date'])
# %%

df.loc[(df.Date >= year_begin) & (df.Date <= today),'ToDate1'] = 'YTD'
# %%
df.loc[(df.Date >= month_begin) & (df.Date <= today),'ToDate2'] = 'MTD'

# %%
df.loc[(df.Date >= month_begin) & (df.Date <= today),'YearMonth'].unique()

# %%
df = df[df['ToDate1'].notna()].copy()
# %%
df[['Date', 'ToDate1', 'ToDate2']].to_csv(r'\\tpfile01\8200_Accounting\PowerBI\Mapping_General\calendar_to_date.csv', index=False)
