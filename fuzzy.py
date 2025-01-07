'''
@-*- coding: utf-8 -*-
@Project ：textmatch
@File    ：fuzzy.py
@Author  ：Weaud
@Date    ：2024/12/25
@explain : 文件说明
'''
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

# 模糊匹配

def fuzzy_merge(df_1, df_2, key1, key2, threshold, limit=2):
    """
    :param df_1: the left table to join
    :param df_2: the right table to join
    :param key1: key column of the left table
    :param key2: key column of the right table
    :param threshold: how close the matches should be to return a match, based on Levenshtein distance
    :param limit: the amount of matches that will get returned, these are sorted high to low
    :return: dataframe with boths keys and matches
    """
    s = df_2[key2].tolist()

    m = df_1[key1].apply(lambda x: process.extract(x, s, limit=limit))
    df_1['matches'] = m

    m2 = df_1['matches'].apply(
        lambda x: [i[0] for i in x if i[1] >= threshold][0] if len([i[0] for i in x if i[1] >= threshold]) > 0 else '')
    df_1['matches'] = m2

    return df_1

dengji = pd.read_excel('dengji.xlsx')
zhuce = pd.read_excel('zhuce.xlsx')

df = fuzzy_merge(dengji, zhuce, '登记名称', '注册名称', threshold=10)
df.to_excel('matched_data_fuzz.xlsx', index=False)