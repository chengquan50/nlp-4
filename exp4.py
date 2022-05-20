# coding=gbk
import jieba
import gensim
import openpyxl
from gensim.models import word2vec
from gensim.models import Word2Vec
# def load_stopword():
#     '''
#     加载停用词表
#     :return: 返回停用词的列表
#     '''
#     f_stop = open('cn_stopwords.txt', encoding='utf-8')
#     sw = [line.strip() for line in f_stop]
#     f_stop.close()
#     return sw
# def load_text(stop_words,f):
#     texts = []
#     for line in f:
#         cd = jieba.lcut(line)
#         cddd = []
#         for dd in cd:
#             if dd not in stop_words:
#                 cddd.append(dd)
#         texts.append(cddd)
#     return texts
# # 加载停用词表
# stop_words = load_stopword()
# # 读入语料库
# f = open('preprocess.txt', "r", encoding='utf-8')
# texts=load_text(stop_words,f)
# f.close()
# file = open('op.txt','w',encoding='utf-8')
# for line in texts:
#     for w in line:
#         file.write(w+' ')
#     file.write('\n')
# sentences=word2vec.Text8Corpus('op.txt')
# model=word2vec.Word2Vec(sentences,sg=0,window=5,min_count=2,negative=3,sample=0.001,hs=1,workers=4)
# y1=model.wv.most_similar("郭靖",topn=10)
# print(y1)
# y2=model.wv.most_similar("黄蓉",topn=10)
# print(y2)
# y3=model.wv.most_similar("郭靖",topn=10)
# print(y3)
# y4=model.wv.most_similar("黄蓉",topn=10)
# print(y4)
from gensim.models import KeyedVectors
sentences=word2vec.Text8Corpus('op.txt')
model=word2vec.Word2Vec(sentences,sg=1,window=5,min_count=2,negative=3,sample=0.001,hs=1,workers=4)
word_vectors = model.wv
word_vectors.save("word2vec1.wordvectors")
wv = KeyedVectors.load("word2vec1.wordvectors", mmap='r')
# vector = wv['郭靖']  # Get numpy vector of a word
with open("results.txt",'a') as f:
    print("袁承志",wv.most_similar(u"袁承志", topn=10),file=f)
    print("黄蓉",wv.most_similar(u"黄蓉", topn=10),file=f)
    print("杨过",wv.most_similar(u"杨过", topn=10),file=f)
    print("令狐冲",wv.most_similar(u"令狐冲", topn=10),file=f)
    print("韦小宝",wv.most_similar(u"韦小宝", topn=10),file=f)
    print("峨嵋派",wv.most_similar(u"峨嵋派", topn=10),file=f)
    print("葵花宝典",wv.most_similar(u"葵花宝典", topn=10),file=f)
    print("九阴真经",wv.most_similar(u"九阴真经", topn=10),file=f)
vc=["郭靖","黄药师","黄蓉","杨康","穆念慈","欧阳锋","段智兴","周伯通","梅超风","柯镇恶","欧阳克","王重阳","完颜洪烈","郭啸天","包惜弱"]
tian=["段正明","段正淳","段延庆","刀白凤","秦红棉","甘宝宝","阮星竹","王语嫣","阿朱","阿紫",'慕容复','鸠摩智']
shen=["一灯大师","小龙女","尹志平","丘处机","公孙止","孙婆婆","李莫愁","李志常","完颜萍","陆无双","杨过","金轮法王","洪七公","郭靖","郭襄"]
xiao=["东方不败","风清扬","方证大师","任我行","令狐冲","冲虚道长","岳不群","林平之","左冷禅","解风","向问天"]
vc+=tian
vc+=shen
vc+=xiao
wb = openpyxl.load_workbook("Test.xlsx")
sheet = wb['Sheet1']
for eve in vc:
    name=[eve]
    name+=list(wv[eve])
    sheet.append(name)
wb.save("Test.xlsx")

# model ['黄蓉']