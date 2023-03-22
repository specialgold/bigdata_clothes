import pandas as pd
import numpy as np
from mlxtend.preprocessing import TransactionEncoder
from mlxtend.frequent_patterns import apriori, association_rules
def aggregation(df):
    # df['site']=df['관리번호(사이트-YYYY-MM-W-NN)'].str[:2]
    df['ym'] = df['관리번호(사이트-YYYY-MM-W-NN)'].str[3:11]
    # df['y'] = df['관리번호(사이트-YYYY-MM-W-NN)'].str[3:7]
    # df['m'] = df['관리번호(사이트-YYYY-MM-W-NN)'].str[8:11]
    # print(df.head(3))

    # print(df.groupby('ym')['카테고리'].value_counts())

    # print(df.groupby(['ym', '카테고리'])['기장'].value_counts())

    # print(df.groupby(['ym', '카테고리', '기장'])['네크라인'].value_counts())

    # print(df.groupby(['ym', '아이템', '기장','네크라인','소매기장','핏','셰이프','숄더','슬리브'])['웨이스트라인'].value_counts())
    result = df.groupby(['ym', '아이템', '기장', '네크라인', '소매기장', '핏', '셰이프', '숄더'])['슬리브'].value_counts()
    print(result)
    result.to_excel('groupby_result.xlsx')
    # news = result.reset_index()
    # news.to_excel('groupby_result.xlsx')
    # print(df.groupby(['ym', '아이템', '기장', '네크라인', '소매기장'])['핏'].value_counts())

    # print(df.groupby('아이템')['기장'].value_counts())

def frequentItem():
    for file in ["2020", "2021", "2022"]:
    #for file in ["2020"]:
        df = pd.read_excel(file+".xlsx")
        transactions = []
        for i in range(len(df)):
            line = df.loc[i, "디테일"]
            trans = []
            items = line.split(',')
            for item in items:
                if len(item) != 0:
                    trans.append(item)
            transactions.append(trans)
            #print(line)
        # TransactionEncoder를 사용하여 데이터를 변환합니다.

        te = TransactionEncoder()
        te_ary = te.fit(transactions).transform(transactions)
        df = pd.DataFrame(te_ary, columns=te.columns_)

        # 지지도(support)가 0.01 이상인 항목 집합을 찾습니다.
        frequent_itemsets = apriori(df, min_support=0.01, use_colnames=True)

        # 신뢰도(confidence)가 0.5 이상인 연관 규칙을 찾습니다.
        rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=0.5)

        # 결과를 출력합니다.
        print(file)
        print(rules)



"""

    df['ym'] = df['관리번호(사이트-YYYY-MM-W-NN)'].str[3:11]
"""

"""
import pandas as pd
from mlxtend.preprocessing import TransactionEncoder
from mlxtend.frequent_patterns import apriori, association_rules

# 데이터를 불러옵니다.
df = pd.read_excel("2020.xlsx")

# 각 행을 리스트로 변환합니다.
transactions = []
for i in range(df.shape[0]):
    transactions.append([str(df.values[i, j]) for j in range(df.shape[1])])

# TransactionEncoder를 사용하여 데이터를 변환합니다.
te = TransactionEncoder()
te_ary = te.fit(transactions).transform(transactions)
df = pd.DataFrame(te_ary, columns=te.columns_)

# 지지도(support)가 0.01 이상인 항목 집합을 찾습니다.
frequent_itemsets = apriori(df, min_support=0.01, use_colnames=True)

# 신뢰도(confidence)가 0.5 이상인 연관 규칙을 찾습니다.
rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=0.5)

# 결과를 출력합니다.
print(rules)
"""
# df = pd.read_excel("bestselling.xlsx")
# print(df)
# print(df.head(3))
# aggregation(df)
frequentItem()



