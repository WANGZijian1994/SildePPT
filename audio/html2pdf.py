import os
import asyncio
from playwright.async_api import async_playwright

async def html_to_pdf_folder(folder_path):
    async with async_playwright() as p:
        # 启动浏览器
        browser = await p.chromium.launch()
        page = await browser.new_page()

        # 遍历文件夹
        for fname in os.listdir(folder_path):
            if fname.endswith('.html'):
                html_path = os.path.join(folder_path, fname)
                # 转换输出路径：a.html -> a.pdf
                pdf_path = html_path.rsplit('.', 1)[0] + '.pdf'
                
                # 关键：将本地路径转为浏览器能认的 file:/// 格式
                file_url = 'file:///' + os.path.abspath(html_path).replace('\\', '/')
                
                print(f"正在转换: {fname}...")
                
            try:
                # 像浏览器一样打开文件
                await page.goto(file_url)
                
                # --- 修正这里：正确的方法名 ---
                await page.wait_for_load_state('networkidle') 
                
                # 打印成 PDF
                await page.pdf(path=pdf_path, format="A4", print_background=True)
                print(f"✅ 成功: {pdf_path}")
            except Exception as e:
                print(f"❌ 失败 {fname}: {e}")

    await browser.close()

# 你的文件夹路径
folder = r'D:\\副业赚钱\\教会事务\\SildePPT\\audio'
asyncio.run(html_to_pdf_folder(folder))