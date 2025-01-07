'''
@-*- coding: utf-8 -*-
@Project ：fuzzy.py 
@File    ：fuzz_stop.py
@Author  ：Weaud
@Date    ：2024/12/27
@explain : 文件说明
'''
import pandas as pd
import jieba
from fuzzywuzzy import fuzz

# 加载数据
dengji = pd.read_excel('dengji.xlsx')
zhuce = pd.read_excel('zhuce.xlsx')

# 假设中文名称字段为 "名称"
dengji_names = dengji['登记名称'].tolist()
zhuce_names = zhuce['注册名称'].tolist()

# 高频词列表
stop_words = ['股份', '有限', '公司', '集团', '安徽', '合肥']


# 定义一个函数，用 jieba 分词并去除高频词
def simplify_name(name, stop_words):
    words = jieba.cut(name)
    simplified = ''.join([word for word in words if word not in stop_words])
    return simplified


# 对两个列表的名称进行简化处理
simplified_dengji_names = [simplify_name(name, stop_words) for name in dengji_names]
simplified_zhuce_names = [simplify_name(name, stop_words) for name in zhuce_names]

# 创建一个匹配结果列表
matched_results = []

# 对每个登记名称，找到最相似的注册名称
for dengji_name, simplified_dengji_name in zip(dengji_names, simplified_dengji_names):
    best_match = None
    best_score = -1

    for zhuce_name, simplified_zhuce_name in zip(zhuce_names, simplified_zhuce_names):
        # 计算相似度
        score = fuzz.ratio(simplified_dengji_name, simplified_zhuce_name)

        # 更新最佳匹配
        if score > best_score:
            best_score = score
            best_match = zhuce_name

    # 将结果保存到列表
    matched_results.append([dengji_name, best_match, best_score])

# 将结果保存到新的 Excel 文件中
matched_df = pd.DataFrame(matched_results, columns=['登记名称', '匹配注册名称', '相似度'])
matched_df.to_excel('matched_data_fuzz_stop.xlsx', index=False)

print("匹配结果已保存到 'matched_data_fuzz_stop.xlsx' 文件中。")
