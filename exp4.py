# coding=gbk
import jieba
import gensim
import openpyxl
from gensim.models import word2vec
from gensim.models import Word2Vec
# def load_stopword():
#     '''
#     ����ͣ�ôʱ�
#     :return: ����ͣ�ôʵ��б�
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
# # ����ͣ�ôʱ�
# stop_words = load_stopword()
# # �������Ͽ�
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
# y1=model.wv.most_similar("����",topn=10)
# print(y1)
# y2=model.wv.most_similar("����",topn=10)
# print(y2)
# y3=model.wv.most_similar("����",topn=10)
# print(y3)
# y4=model.wv.most_similar("����",topn=10)
# print(y4)
from gensim.models import KeyedVectors
sentences=word2vec.Text8Corpus('op.txt')
model=word2vec.Word2Vec(sentences,sg=1,window=5,min_count=2,negative=3,sample=0.001,hs=1,workers=4)
word_vectors = model.wv
word_vectors.save("word2vec1.wordvectors")
wv = KeyedVectors.load("word2vec1.wordvectors", mmap='r')
# vector = wv['����']  # Get numpy vector of a word
with open("results.txt",'a') as f:
    print("Ԭ��־",wv.most_similar(u"Ԭ��־", topn=10),file=f)
    print("����",wv.most_similar(u"����", topn=10),file=f)
    print("���",wv.most_similar(u"���", topn=10),file=f)
    print("�����",wv.most_similar(u"�����", topn=10),file=f)
    print("ΤС��",wv.most_similar(u"ΤС��", topn=10),file=f)
    print("������",wv.most_similar(u"������", topn=10),file=f)
    print("��������",wv.most_similar(u"��������", topn=10),file=f)
    print("�����澭",wv.most_similar(u"�����澭", topn=10),file=f)
vc=["����","��ҩʦ","����","�","�����","ŷ����","������","�ܲ�ͨ","÷����","�����","ŷ����","������","���պ���","��Х��","��ϧ��"]
tian=["������","������","������","���׷�","�غ���","�ʱ���","������","������","����","����",'Ľ�ݸ�','�Ħ��']
shen=["һ�ƴ�ʦ","С��Ů","��־ƽ","�𴦻�","����ֹ","������","��Ī��","��־��","����Ƽ","½��˫","���","���ַ���","���߹�","����","����"]
xiao=["��������","������","��֤��ʦ","������","�����","�������","����Ⱥ","��ƽ֮","������","���","������"]
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

# model ['����']