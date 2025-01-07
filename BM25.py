import pandas as pd
from rank_bm25 import BM25Okapi
import jieba  # 用于中文分词

# 加载数据
dengji = pd.read_excel('dengji.xlsx')
zhuce = pd.read_excel('zhuce.xlsx')

# 指定需要匹配的字段
dengji_column = '登记名称'
zhuce_column = '注册名称'

# 对数据中的中文名称进行分词
dengji['分词'] = dengji[dengji_column].apply(lambda x: list(jieba.cut(str(x))))
zhuce['分词'] = zhuce[zhuce_column].apply(lambda x: list(jieba.cut(str(x))))

# 构建 BM25 模型
bm25 = BM25Okapi(zhuce['分词'].tolist())

# 进行匹配
def match_name(row):
    query = row['分词']
    scores = bm25.get_scores(query)
    best_idx = scores.argmax()  # 找到得分最高的索引
    best_match = zhuce.iloc[best_idx][zhuce_column]
    best_score = scores[best_idx]

    # 只保留得分大于 5 的匹配
    if best_score > 0:
        return pd.Series([best_match, best_score])
    else:
        return pd.Series([None, None])  # 不符合条件的置为空

# 应用匹配函数
dengji[['匹配的注册名称', '匹配得分']] = dengji.apply(match_name, axis=1)

# 保存结果到新的 Excel 文件
output_path = 'matched_data_bm25.xlsx'
dengji.to_excel(output_path, index=False)
print(f"匹配结果已保存到 {output_path}")
