from fuzzychinese import FuzzyChineseMatch
from fuzzywuzzy import process
import re  # 正则表达式


# import pandas as pd

# 定义目标词汇列表
TARGET_WORDS = ['工程或费用名称', '建筑工程', '安装工程', '设备及工器具购置费', '其他费用', '合计', '单位', '数量',
                '单位价值（元）', '备注',
                '项', '目', '节', '细目', '序号']
Atlas_PipeFittingsQ235A = {}  # 管配件重量表
Atlas_Valve = {}  # 阀门价格表
def init_atlas(pAtlas_PipeFittingsQ235A, pAtlas_valve):
    global Atlas_PipeFittingsQ235A, Atlas_Valve
    Atlas_PipeFittingsQ235A = pAtlas_PipeFittingsQ235A
    Atlas_Valve = pAtlas_valve
    print(Atlas_PipeFittingsQ235A.keys())



def fuzzy_match(p_raw_word, p_target_words=None):
    # 初始化 FuzzyChineseMatch 对象
    if p_target_words is None:
        p_target_words = TARGET_WORDS
    _fcm = FuzzyChineseMatch(ngram_range=(3, 3), analyzer='stroke')

    p_raw_word = re.sub(r'[^\u4e00-\u9fff]+', "", str(p_raw_word))  # 用正则表达式删除字符串中所有非汉字字符，提高识别效率

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
        similarity_scores = _fcm.get_similarity_score()[0]  # 获取相似度分数

        return matched_words, similarity_scores
    else:
        return None, None


def test_para(para):
    para = {'细目', '序号'}
    rtp = para
    return rtp


material_fittings = ["Q235A", "SS304", "橡胶", "PE100", "PVC-U", "nan"]
material_type = ["管配件", "阀门", "设备", "材料"]


def fuzzy_match_EM(pEquipmentMaterial):
    """
    输入一个字符串，返回最匹配的管配件名称及其相似度得分。

    参数:
        input_string (str): 输入的字符串

    返回:
        tuple: (最匹配的管配件名称, 相似度得分)
    """
    # 使用 fuzzywuzzy 的 process.extractOne 方法进行模糊匹配
    best_match_material, score = process.extractOne(pEquipmentMaterial.material, material_fittings)
    rtBest_match = ""
    rtScore = 0
    if best_match_material == "Q235A" or best_match_material == "SS304":
        print(Atlas_PipeFittingsQ235A.keys())
        match_result = process.extractOne(pEquipmentMaterial.name, Atlas_PipeFittingsQ235A.keys())
        if match_result is not None:  # 检查是否找到匹配
            best_match, score = match_result
            print(best_match, score)
            if score > rtScore:
                rtBest_match = best_match
                rtScore = score

    return rtBest_match, rtScore


def extract_specifications(spec_string):
    """
    从规格字符串中提取管径和长度参数。
    参数:
        spec_string (str): 规格字符串，例如 "DN1=1200,DN2=500,DN3=500" 或 "DN1400，L=9000"
    返回:
        dict: 包含提取的管径和长度参数
    """
    # 初始化结果字典
    result = {"管径": [], "长度": 0, "单位": ""}

    # 提取管径
    # 匹配模式：DN后可能跟标识符（如DN1），然后是等号和数字
    diameter_pattern = re.compile(r'DN\d*\s*=\s*(\d+)|DN(\d+)')
    diameter_matches = diameter_pattern.findall(spec_string)
    for match in diameter_matches:
        for diameter in match:
            if diameter:
                result["管径"].append(int(diameter))
    result["管径"].sort(reverse=True)

    # 提取长度
    # 匹配模式：L或La后跟全角或半角等号和数字，可能有单位
    length_pattern = re.compile(r'(L|La)\s*[=＝]\s*(\d+)(mm|cm|m)?')
    length_match = length_pattern.search(spec_string)
    if length_match:
        length_value = int(length_match.group(2))
        length_unit = length_match.group(3) if length_match.group(3) else "mm"  # 默认单位为 mm
        result["长度"] = float(length_value)
        result["单位"] = length_unit

    return result
