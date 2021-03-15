import pandas as pd
pd.set_option('display.max_rows',None)
# 读取数据
# test_df = pd.read_csv("./data/test.csv")
train_df = pd.read_csv("./zhangdie.csv")

# train_df.head()
# print(train_df.head())
# print(train_df.tail())
# print(train_df.info())

# # 对数值型变量描述
print(train_df.describe())
# # 对标称型变量描述
print(train_df.describe(include=['O']))  

# # 查看最有影响的几个值，分别对存活率的影响
机构 = train_df[['机构','研报前后涨跌幅']].groupby('机构').mean().sort_values(by='研报前后涨跌幅')
评级 = train_df[['原文评级','研报前后涨跌幅']].groupby('原文评级').mean()
变动 = train_df[{'评级变动','研报前后涨跌幅'}].groupby('评级变动').mean()

print(机构)
print(评级)
print(变动)
# print(train_df['评级变动'])

from scipy.stats import ttest_ind
tiaogao=train_df[train_df['评级变动']=='调高']
weichi=train_df[train_df['评级变动']=='维持']
tiaodi=train_df[train_df['评级变动']=='调低']
print(ttest_ind(tiaogao['研报前后涨跌幅'],tiaodi['研报前后涨跌幅']))