from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# 注册中文字体
try:
    # Windows 系统字体路径
    pdfmetrics.registerFont(TTFont('SimSun', 'C:/Windows/Fonts/simsun.ttc'))
    pdfmetrics.registerFont(TTFont('SimHei', 'C:/Windows/Fonts/simhei.ttf'))
    FONT_NAME = 'SimSun'
    FONT_BOLD = 'SimHei'
except:
    print("警告：无法加载中文字体，尝试使用备用字体")
    try:
        pdfmetrics.registerFont(TTFont('SimSun', 'simsun.ttc'))
        pdfmetrics.registerFont(TTFont('SimHei', 'simhei.ttf'))
        FONT_NAME = 'SimSun'
        FONT_BOLD = 'SimHei'
    except:
        print("错误：无法加载中文字体，将使用默认字体（中文可能无法正常显示）")
        FONT_NAME = 'Helvetica'
        FONT_BOLD = 'Helvetica-Bold'

def draw_logo(c, page_width, page_height, logo_path="爱心教会.png"):
    """在右上角绘制logo"""
    try:
        logo_width = 80  # logo宽度
        logo_height = 80  # logo高度
        x_position = page_width - logo_width - 20  # 距离右边20像素
        y_position = page_height - logo_height - 20  # 距离顶部20像素
        
        c.drawImage(logo_path, x_position, y_position,
                    width=logo_width, height=logo_height,
                    preserveAspectRatio=True, mask='auto')
    except Exception as e:
        print(f"警告：无法加载logo图片 {logo_path}: {e}")

def add_content_with_auto_paging(c, content_lines, page_width, page_height):
    """自动分页添加内容"""
    margin_top = 150  # 增加上边距，留更多白
    margin_left = 60
    margin_right = 60
    margin_bottom = 80
    y_position = page_height - margin_top
    max_width = page_width - margin_left - margin_right
    
    for line in content_lines:
        # 跳过空行但保留一定间距
        if not line.strip():
            y_position -= 25
            continue
        
        # 确定字体和缩进 - 投影用，字体加大
        if line.strip().startswith(("一，", "二，", "三，", "四，", "五，")):
            # 主标题
            font_name = FONT_BOLD
            font_size = 32
            line_height = 46
            indent = margin_left
        elif "祷告事项" in line or "开场白" in line:
            # 特殊标题
            font_name = FONT_BOLD
            font_size = 30
            line_height = 44
            indent = margin_left
        elif line.strip() and line.strip()[0].isdigit() and "," in line[:5]:
            # 小标题（如"1,为..."）
            font_name = FONT_BOLD
            font_size = 26
            line_height = 38
            indent = margin_left + 20
        else:
            # 普通内容 - 改为加粗
            font_name = FONT_BOLD
            font_size = 24
            line_height = 36
            indent = margin_left + 30
        
        c.setFont(font_name, font_size)
        
        # 自动换行
        if c.stringWidth(line, font_name, font_size) > max_width - (indent - margin_left):
            words = line
            temp_line = ""
            for char in words:
                test_line = temp_line + char
                if c.stringWidth(test_line, font_name, font_size) < max_width - (indent - margin_left):
                    temp_line = test_line
                else:
                    # 检查是否需要换页
                    if y_position < margin_bottom:
                        c.showPage()
                        draw_logo(c, page_width, page_height)
                        y_position = page_height - margin_top
                    
                    c.drawString(indent, y_position, temp_line)
                    y_position -= line_height
                    temp_line = char
            
            # 写入最后一行
            if temp_line:
                if y_position < margin_bottom:
                    c.showPage()
                    draw_logo(c, page_width, page_height)
                    y_position = page_height - margin_top
                c.drawString(indent, y_position, temp_line)
                y_position -= line_height
        else:
            # 检查是否需要换页
            if y_position < margin_bottom:
                c.showPage()
                draw_logo(c, page_width, page_height)
                y_position = page_height - margin_top
            
            c.drawString(indent, y_position, line)
            y_position -= line_height
    
    return y_position

def generate_prayer_meeting_pdf(txt_filename="./新年祷告会.txt", pdf_filename="2026年1月新年祷告会.pdf"):
    """生成新年祷告会PDF"""
    # 创建PDF文件
    c = canvas.Canvas(pdf_filename, pagesize=A4)
    width, height = A4
    
    # 页面1：封面
    # 绘制红色背景
    c.setFillColorRGB(0.8, 0, 0)  # 红色背景
    c.rect(0, 0, width, height, fill=1, stroke=0)
    
    # 恢复白色用于文字
    c.setFillColorRGB(1, 1, 1)  # 白色文字
    
    draw_logo(c, width, height)
    c.setFont(FONT_BOLD, 50)
    title = "辞旧迎新祷告会"
    title_width = c.stringWidth(title, FONT_BOLD, 50)
    c.drawString((width - title_width) / 2, height / 2 + 80, title)
    
    c.setFont(FONT_NAME, 30)
    subtitle = "2026-01-01"
    subtitle_width = c.stringWidth(subtitle, FONT_NAME, 30)
    c.drawString((width - subtitle_width) / 2, height / 2 + 20, subtitle)
    
    c.setFont(FONT_NAME, 24)
    church = "巴黎基督国度爱心教会"
    church_width = c.stringWidth(church, FONT_NAME, 24)
    c.drawString((width - church_width) / 2, height / 2 - 50, church)
    
    # 恢复黑色用于后续页面
    c.setFillColorRGB(0, 0, 0)
    
    # 从第二页开始，加载txt文件（0.txt到5.txt）
    for i in range(6):
        txt_file = f"{i}.txt"
        
        # 创建新页面
        c.showPage()
        draw_logo(c, width, height)
        
        # 读取txt文件
        if os.path.exists(txt_file):
            with open(txt_file, 'r', encoding='utf-8') as f:
                content = f.read()
            paragraphs = content.split('\n')
            add_content_with_auto_paging(c, paragraphs, width, height)
        else:
            print(f"警告：找不到文件 {txt_file}")
            # 显示提示信息
            c.setFont(FONT_NAME, 16)
            c.drawString(60, height - 100, f"文件 {txt_file} 未找到")
    
    # 保存PDF
    c.save()
    print(f"PDF文件已生成: {pdf_filename}")

if __name__ == "__main__":
    pdf_filename = "D:\\副业赚钱\\教会事务\\PPT\\2026年1月新年祷告会.pdf"
    txt_filename = "D:\\副业赚钱\\教会事务\\PPT\\新年祷告会.txt"
    generate_prayer_meeting_pdf(txt_filename=txt_filename, pdf_filename=pdf_filename)