from fuzzychinese import FuzzyChineseMatch
import pandas as pd


def fuzzy_match(raw_word, target_words):
    # 初始化 FuzzyChineseMatch 对象
    fcm = FuzzyChineseMatch(ngram_range=(3, 3), analyzer='stroke')

    # 训练模型
    fcm.fit(target_words)

    # 使用 transform 方法查找与 raw_word 最相近的词
    matches = fcm.transform([raw_word], n=1)

    # 检查返回的匹配结果是否包含足够的值
    if matches and len(matches[0]) == 2:
        matched_word, similarity_score = matches[0]
    else:
        matched_word, similarity_score = None, None

    # 返回匹配结果
    return matched_word, similarity_score


