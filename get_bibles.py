import requests
from urllib.parse import quote
from zhconv import convert

# https://bible-api.com/%E8%B7%AF%E5%8A%A0%E7%A6%8F%E9%9F%B3+1:27?translation=cuv

def get_bible_verses(book_name, chapter, start_verse, end_verse, French=False):
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
    url = f"https://bible-api.com/{encoded_book}+{chapter}:{start_verse}-{end_verse}?translation=cuv"
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
book = "路加福音"      # 路加福音
chapter_num = 9    # 第1章
start = 1          # 第1节
end = 27            # 到第5节

result = get_bible_verses(book, chapter_num, start, end)

# 打印结果
for i, text in enumerate(result, start=start):
    print(f"第{i}节: {text}")
'''