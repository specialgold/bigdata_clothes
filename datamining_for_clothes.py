import pandas as pd
import numpy as np

df = pd.read_excel("bestselling.xlsx")
# print(df)

print(df.head(3))

# df['site']=df['관리번호(사이트-YYYY-MM-W-NN)'].str[:2]
df['ym'] = df['관리번호(사이트-YYYY-MM-W-NN)'].str[3:11]
# df['y'] = df['관리번호(사이트-YYYY-MM-W-NN)'].str[3:7]
# df['m'] = df['관리번호(사이트-YYYY-MM-W-NN)'].str[8:11]
# print(df.head(3))

# print(df.groupby('ym')['카테고리'].value_counts())

# print(df.groupby(['ym', '카테고리'])['기장'].value_counts())

# print(df.groupby(['ym', '카테고리', '기장'])['네크라인'].value_counts())



print(df.groupby('아이템')['기장'].value_counts())
