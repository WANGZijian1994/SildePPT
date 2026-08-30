import requests
from urllib.parse import quote
from zhconv import convert

indexes = {
    "约翰福音": "約翰福音",
    "传道书": "ecc",
    "路加福音": "路加福音",
    "撒母耳记上": "1SA",
    "雅各书": "JAS",
    "路得记": "RUT",
    "哥林多前书": "1CO",
    "提多书": "TIT",
    "提摩太前书": "1TI",
    "提摩太后书": "2TI",
    "马太福音": "MAT",
    "以赛亚书": "ISA",
    "箴言": "PRO",
    "诗篇": "PSA",
    "帖撒罗尼迦前书": "1TH",
    "帖撒罗尼迦后书": "2TH",
    "马可福音": "MRK",
    "腓立比书": "PHP",
    "约翰一书": "1JN",
    "约翰二书": "2JN",
    "启示录": "REV",
    "歌罗西书": "COL",
    "腓利门书": "PHM",
    "使徒行传": "ACT",
    "以弗所书": "EPH",
}

# https://bible-api.com/%E8%B7%AF%E5%8A%A0%E7%A6%8F%E9%9F%B3+1:27?translation=cuv

def get_bible_verses(book_name, chapter, start_verse, end_verse, French=False, version='cuv'):
    """
    获取指定章节和范围的简体中文经文
    :param book_name: 圣经卷名 (中文或英文标识，如 "路加福音" 或 "Luke")
    :param chapter: 第几章
    :param start_verse: 起始节
    :param end_verse: 结束节
    :return: 经文列表
    """
    # 使用 Bible-api，指定版本为 cuv (和合本简体)
    # 格式：https://bible-api.com/book+chapter:start-end?translation=cuv
    # 对中文书名进行 URL 编码
    encoded_book = quote(book_name)
    url = f"https://bible-api.com/{encoded_book}+{chapter}:{start_verse}-{end_verse}?translation={version}"
    if French:
        url = f"https://bible-api.com/{encoded_book}+{chapter}:{start_verse}-{end_verse}?translation=lsf"  # 法语版本
    
    try:
        response = requests.get(url)
        response.raise_for_status() # 检查请求是否成功
        data = response.json()
        
        # 提取每一节的内容并存入列表
        verses_list = [convert(verse['text'].strip(), 'zh-cn') for verse in data['verses']]
        return verses_list

    except Exception as e:
        return [f"错误: 无法获取数据 ({e})"]

# --- 使用示例 ---
# 常见的英文对应：路加福音 -> Luke, 创世记 -> Genesis, 马太福音 -> Matthew

'''
book = "约翰福音"      # 路加福音 約翰福音 ecc 传道书
chapter_num = 1    # 第1章
start = 14         # 第1节
end = 15           # 到第5节

result = get_bible_verses(indexes[book], chapter_num, start, end)

# 打印结果
for i, text in enumerate(result, start=start):
    print(f"第{i}节: {text}")
'''

