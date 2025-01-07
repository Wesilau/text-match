'''
@-*- coding: utf-8 -*-
@Project ：textmatch 
@File    ：similarity.py
@Author  ：Weaud
@Date    ：2024/12/25
@explain : 文件说明
'''
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import numpy as np
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# 加载数据
dengji = pd.read_excel('dengji.xlsx')
zhuce = pd.read_excel('zhuce.xlsx')

# 提取中文名称字段，假设字段名为'名称'
dengji_names = dengji['登记名称'].tolist()
zhuce_names = zhuce['注册名称'].tolist()


# 计算 Levenshtein 距离（编辑距离）相似度
def levenshtein_similarity(name1, name2):
    return fuzz.ratio(name1, name2) / 100.0


# 计算 Jaccard 相似度
def jaccard_similarity(name1, name2):
    set1 = set(name1)
    set2 = set(name2)
    intersection = len(set1.intersection(set2))
    union = len(set1.union(set2))
    return intersection / union


# 计算 Cosine 相似度
def cosine_similarity_fn(name1, name2):
    vectorizer = CountVectorizer().fit_transform([name1, name2])
    cosine_sim = cosine_similarity(vectorizer[0:1], vectorizer[1:2])
    return cosine_sim[0][0]


# 创建一个新的匹配结果列表
matched_results = []

# 对于每个登记名称，找到最佳匹配的注册名称
for name1 in dengji_names:
    best_match = None
    best_score_lev = -1
    best_score_jaccard = -1
    best_score_cosine = -1
    best_match_lev = None
    best_match_jaccard = None
    best_match_cosine = None

    for name2 in zhuce_names:
        # 计算 Levenshtein 相似度
        score_lev = levenshtein_similarity(name1, name2)
        if score_lev > best_score_lev:
            best_score_lev = score_lev
            best_match_lev = name2

        # 计算 Jaccard 相似度
        score_jaccard = jaccard_similarity(name1, name2)
        if score_jaccard > best_score_jaccard:
            best_score_jaccard = score_jaccard
            best_match_jaccard = name2

        # 计算 Cosine 相似度
        score_cosine = cosine_similarity_fn(name1, name2)
        if score_cosine > best_score_cosine:
            best_score_cosine = score_cosine
            best_match_cosine = name2

    matched_results.append(
        [name1, best_match_lev, best_score_lev, best_match_jaccard, best_score_jaccard, best_match_cosine,
         best_score_cosine])

# 将结果保存到新的 Excel 文件中
matched_df = pd.DataFrame(matched_results,
                          columns=['dengji', '最佳匹配_Levenshtein', 'L相似度', '最佳匹配_Jaccard',
                                   'J相似度', '最佳匹配_Cosine', 'C相似度'])
matched_df.to_excel('matched_data_sim.xlsx', index=False)

print("匹配结果已保存到 'matched_data_sim.xlsx' 文件中。")
