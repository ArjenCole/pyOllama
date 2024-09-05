from fuzzychinese import FuzzyChineseMatch
import pandas as pd


def fuzzy_match(p_raw_word):
    # 初始化 FuzzyChineseMatch 对象
    _fcm = FuzzyChineseMatch(ngram_range=(3, 3), analyzer='stroke')

    # 定义目标词汇列表
    _target_words = ['工程或费用名称', '建筑工程', '安装工程', '设备及工器具购置费', '其他费用', '合计', '单位', '数量',
                     '单位价值（元）', '备注',
                     '项', '目', '节', '细目', '序号']

    # 训练模型
    _fcm.fit(_target_words)

    # 使用 transform 方法查找与 p_raw_word 最相近的词
    rt_matches = _fcm.transform([p_raw_word], n=3)

    print('rt_matches', rt_matches)
    # 获取匹配的字符串及其相似度分数
    # 检查是否有匹配结果
    if 1 == 1:
        matched_words = rt_matches[0]  # 这里假设我们只关心第一个匹配结果
        similarity_scores = _fcm.get_similarity_score()  # 获取相似度分数
        return matched_words, similarity_scores
    else:
        return None, None

