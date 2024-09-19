from fuzzychinese import FuzzyChineseMatch
import pandas as pd

# 定义目标词汇列表
TARGET_WORDS = ['工程或费用名称', '建筑工程', '安装工程', '设备及工器具购置费', '其他费用', '合计', '单位', '数量',
                '单位价值（元）', '备注',
                '项', '目', '节', '细目', '序号']


def fuzzy_match(p_raw_word, p_target_words=None):
    # 初始化 FuzzyChineseMatch 对象
    if p_target_words is None:
        p_target_words = TARGET_WORDS
    _fcm = FuzzyChineseMatch(ngram_range=(3, 3), analyzer='stroke')

    # 训练模型
    _fcm.fit(p_target_words)

    try:
        # 尝试执行可能引发异常的代码
        # 使用 transform 方法查找与 p_raw_word 最相近的词
        rt_matches = _fcm.transform([p_raw_word], n=3)
    except:
        # 处理异常
        rt_matches = None

    # print('rt_matches', rt_matches)
    # 获取匹配的字符串及其相似度分数
    # 检查是否有匹配结果
    if rt_matches is not None:
        matched_words = rt_matches[0]  # 这里假设我们只关心第一个匹配结果
        similarity_scores = _fcm.get_similarity_score()  # 获取相似度分数
        return matched_words, similarity_scores
    else:
        return None, None


def test_para(para):
    para = {'细目', '序号'}
    rtp = para
    return rtp
