import jieba
from wordcloud import WordCloud
import pandas as pd
import math
import matplotlib.pyplot as plt


df = pd.read_excel('銷售業績.xlsx')
pivot_df = pd.pivot_table(df,index='產品名稱',
                       values='數量',
                       aggfunc='sum',
                       fill_value=0)

total = pivot_df.sum() #總銷售量33732

pivot_df['銷售數量(%)'] = pivot_df['數量'].map(lambda x:int(math.ceil((x/total)*100))) #新增一欄銷售%，後面當字串重複的次數

#產品名稱清單，去除符號
word_l = []
for item in pivot_df.index:
    word_l.append(item.replace('(','').replace(' ','').replace(')',''))

#把產品名稱建成jieba的字典，自定義字典編碼一定要是utf-8
path = 'productnamedict.txt'
with open(path, 'w', encoding='utf-8') as file :
    for item in word_l:
        file.write(item+'\n')
jieba.load_userdict (path)

#生成文字雲的文字字串
text = ''
for i, item in enumerate(word_l) :
    text += item*pivot_df.iloc[i][1]


cut_text = ' '.join(jieba.cut(text))        # 產生分詞的字串
wd = WordCloud(                             # 建立詞雲物件
    font_path=r"C:\Windows\Fonts\msjhbd.ttc",
    background_color="black",width=1000,height=880, collocations=False).generate(cut_text)

plt.imshow(wd)                              # 由WordCloud物件建立詞雲影像檔
plt.axis("off")                             # 關閉顯示軸線
plt.show()                                  # 顯示詞雲影像檔
