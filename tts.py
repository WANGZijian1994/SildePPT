import asyncio
from edge_tts import Communicate

# 2. 定义你的测试单词（涵盖你 PPT 里的三大规则）
# 格式: {文件名: 单词内容}
test_suite = {
    "rule1_table": "la table",     # e 不发音
    "rule2_paris": "Paris",        # s 睡觉
    "rule2_salut": "Salut",        # t 睡觉
    "rule3_sac": "le sac",         # c 值班
    "rule3_chef": "le chef",       # f 值班
    "rule3_avoir": "avoir",        # r 值班
    "mix_forte": "forte"           # e 唤醒了 t
}

async def generate():
    for name, text in test_suite.items():
        # 使用 HenriNeural，一个非常清晰的法国男声
        communicate = Communicate(text, "fr-FR-HenriNeural")
        await communicate.save(f"audio/{name}.mp3")
        print(f"✅ 已生成: {name}.mp3")

# 3. 执行生成
asyncio.run(generate())