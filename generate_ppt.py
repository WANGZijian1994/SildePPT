from pptx import Presentation
import os

def read_pptx(pptx_file):
    """
    读取现有的PPTX文件并返回所有内容信息
    
    Args:
        pptx_file: PPTX文件路径
    
    Returns:
        dict: 包含PPT所有信息的字典
    """
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return None
    
    prs = Presentation(pptx_file)
    ppt_info = {
        'slide_count': len(prs.slides),
        'slides': []
    }
    
    # 遍历所有幻灯片
    for slide_num, slide in enumerate(prs.slides, 1):
        slide_info = {
            'slide_number': slide_num,
            'title': '',
            'shapes': []
        }
        
        # 获取标题
        if slide.shapes.title:
            slide_info['title'] = slide.shapes.title.text
        
        # 遍历所有形状
        for shape_num, shape in enumerate(slide.shapes):
            shape_info = {
                'shape_number': shape_num,
                'type': str(shape.shape_type),
                'has_text': shape.has_text_frame,
                'text': ''
            }
            
            # 获取文本内容
            if shape.has_text_frame:
                text_parts = []
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text_parts.append(run.text)
                shape_info['text'] = ''.join(text_parts)
            
            slide_info['shapes'].append(shape_info)
        
        ppt_info['slides'].append(slide_info)
    
    return ppt_info


def update_pptx_text(pptx_file, output_file, replacements):
    """
    修改PPTX文件中的文字
    
    Args:
        pptx_file: 原PPTX文件路径
        output_file: 输出PPTX文件路径
        replacements: 字典，格式 {'旧文字': '新文字'}
    """
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return False
    
    prs = Presentation(pptx_file)
    
    # 遍历所有幻灯片
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        # 进行替换
                        for old_text, new_text in replacements.items():
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_text)
    
    # 保存
    prs.save(output_file)
    print(f"PPT文件已保存: {output_file}")
    return True


def print_pptx_info(ppt_info):
    """
    打印PPT信息
    
    Args:
        ppt_info: read_pptx函数返回的信息字典
    """
    if not ppt_info:
        return
    
    print(f"总共 {ppt_info['slide_count']} 页")
    print("=" * 60)
    
    for slide_info in ppt_info['slides']:
        print(f"\n第 {slide_info['slide_number']} 页")
        print(f"标题: {slide_info['title']}")
        print(f"形状数量: {len(slide_info['shapes'])}")       
        for shape_info in slide_info['shapes']:
            
            if shape_info['has_text'] and shape_info['text']:
                print(f"  - 文本: {shape_info['text']}...")


def print_pptx_page(ppt_info, page_number):
    """
    打印PPT信息
    
    Args:
        ppt_info: read_pptx函数返回的信息字典
    """
    if not ppt_info:
        return
    
    print(f"总共 {ppt_info['slide_count']} 页")
    print("=" * 60)
    
    for slide_info in ppt_info['slides']:
        if slide_info['slide_number'] != page_number:
            continue
        print(f"\n第 {slide_info['slide_number']} 页")
        print(f"标题: {slide_info['title']}")
        print(f"形状数量: {len(slide_info['shapes'])}")       
        for shape_info in slide_info['shapes']:
            
            if shape_info['has_text'] and shape_info['text']:
                print(f"  - 文本: {shape_info['text']}...")


def update_slide_text(pptx_file, output_file, slide_number, replacements):
    """
    修改指定页的文字内容
    
    Args:
        pptx_file: 原PPTX文件路径
        output_file: 输出PPTX文件路径
        slide_number: 页码（从1开始）
        replacements: 字典，格式 {'旧文字': '新文字'}
    
    Returns:
        bool: 是否成功
    """
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return False
    
    prs = Presentation(pptx_file)
    
    # 检查页码是否有效
    if slide_number < 1 or slide_number > len(prs.slides):
        print(f"错误：页码 {slide_number} 超出范围（共 {len(prs.slides)} 页）")
        return False
    
    # 获取指定页（索引从0开始）
    slide = prs.slides[slide_number - 1]
    
    # 遍历该页的所有形状
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    # 进行替换
                    for old_text, new_text in replacements.items():
                        #print(f"Before {run.text} → {run.text.replace(old_text, new_text)}")
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)
                        #print(f"After {run.text}")
    
    # 保存
    prs.save(output_file)
    print(f"已修改第 {slide_number} 页，文件已保存: {output_file}")
    return True


def update_multiple_slides(pptx_file, output_file, slide_replacements):
    """
    批量修改多页的文字内容
    
    Args:
        pptx_file: 原PPTX文件路径
        output_file: 输出PPTX文件路径
        slide_replacements: 字典，格式 {页码: {'旧文字': '新文字'}}
        
    Example:
        slide_replacements = {
            1: {'标题': '新标题'},
            2: {'内容': '新内容'},
            3: {'2025': '2026'}
        }
    
    Returns:
        bool: 是否成功
    """
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return False
    
    prs = Presentation(pptx_file)
    
    # 遍历需要修改的页
    for slide_number, replacements in slide_replacements.items():
        # 检查页码是否有效
        if slide_number < 1 or slide_number > len(prs.slides):
            print(f"警告：页码 {slide_number} 超出范围（共 {len(prs.slides)} 页），跳过")
            continue
        
        # 获取指定页
        slide = prs.slides[slide_number - 1]
        
        # 遍历该页的所有形状
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        # 进行替换
                        for old_text, new_text in replacements.items():
                            if old_text in run.text:
                                run.text = run.text.replace(old_text, new_text)
        
        print(f"已修改第 {slide_number} 页")
    
    # 保存
    prs.save(output_file)
    print(f"所有修改完成，文件已保存: {output_file}")
    return True


def delete_slide(pptx_file, output_file, slide_number):
    """
    删除指定页
    
    Args:
        pptx_file: 原PPTX文件路径
        output_file: 输出PPTX文件路径
        slide_number: 要删除的页码（从1开始）
    
    Returns:
        bool: 是否成功
    """
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return False
    
    prs = Presentation(pptx_file)
    
    # 检查页码是否有效
    if slide_number < 1 or slide_number > len(prs.slides):
        print(f"错误：页码 {slide_number} 超出范围（共 {len(prs.slides)} 页）")
        return False
    
    # 获取要删除的幻灯片
    rId = prs.slides._sldIdLst[slide_number - 1].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[slide_number - 1]
    
    # 保存
    prs.save(output_file)
    print(f"已删除第 {slide_number} 页，文件已保存: {output_file}")
    return True


def delete_slides(pptx_file, output_file, slide_numbers):
    """
    批量删除多页
    
    Args:
        pptx_file: 原PPTX文件路径
        output_file: 输出PPTX文件路径
        slide_numbers: 要删除的页码列表（从1开始），如 [2, 5, 7]
    
    Returns:
        bool: 是否成功
    """
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return False
    
    prs = Presentation(pptx_file)
    
    # 从大到小排序，从后往前删除，避免索引变化
    slide_numbers_sorted = sorted(slide_numbers, reverse=True)
    
    for slide_number in slide_numbers_sorted:
        # 检查页码是否有效
        if slide_number < 1 or slide_number > len(prs.slides):
            print(f"警告：页码 {slide_number} 超出范围（共 {len(prs.slides)} 页），跳过")
            continue
        
        # 删除幻灯片
        rId = prs.slides._sldIdLst[slide_number - 1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[slide_number - 1]
        print(f"已删除第 {slide_number} 页")
    
    # 保存
    prs.save(output_file)
    print(f"所有删除完成，文件已保存: {output_file}")
    return True






def duplicate_slides(pptx_file, output_file, slide_numbers):
    """
    批量复制多页并插入到各自后面
    
    Args:
        pptx_file: 原PPTX文件路径
        output_file: 输出PPTX文件路径
        slide_numbers: 要复制的页码列表（从1开始），如 [2, 5]
    
    Returns:
        bool: 是否成功
    """
    import copy
    
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return False
    
    prs = Presentation(pptx_file)
    
    # 从大到小排序，从后往前处理，避免索引变化
    slide_numbers_sorted = sorted(slide_numbers, reverse=True)
    
    for slide_number in slide_numbers_sorted:
        # 检查页码是否有效
        if slide_number < 1 or slide_number > len(prs.slides):
            print(f"警告：页码 {slide_number} 超出范围（共 {len(prs.slides)} 页），跳过")
            continue
        
        # 获取要复制的幻灯片
        source_slide = prs.slides[slide_number - 1]
        
        # 获取布局
        slide_layout = source_slide.slide_layout
        
        # 创建新幻灯片
        new_slide = prs.slides.add_slide(slide_layout)
        
        # 深度复制所有形状
        for shape in source_slide.shapes:
            el = shape.element
            newel = copy.deepcopy(el)
            new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
        
        # 移动到正确位置
        xml_slides = prs.slides._sldIdLst
        slides = list(xml_slides)
        xml_slides.remove(slides[-1])
        xml_slides.insert(slide_number, slides[-1])
        
        print(f"已在第 {slide_number} 页后插入副本")
    
    # 保存
    prs.save(output_file)
    print(f"所有复制完成，文件已保存: {output_file}")
    return True


def duplicate_slide(pptx_file, output_file, slide_number):
    """
    复制指定页并插入到该页后面
    
    Args:
        pptx_file: 原PPTX文件路径
        output_file: 输出PPTX文件路径
        slide_number: 要复制的页码（从1开始）
    
    Returns:
        bool: 是否成功
    """
    import copy
    
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return False
    
    prs = Presentation(pptx_file)
    
    # 检查页码是否有效
    if slide_number < 1 or slide_number > len(prs.slides):
        print(f"错误：页码 {slide_number} 超出范围（共 {len(prs.slides)} 页）")
        return False
    
    # 获取要复制的幻灯片
    source_slide = prs.slides[slide_number - 1]
    
    # 获取布局
    slide_layout = source_slide.slide_layout
    
    # 创建新幻灯片
    new_slide = prs.slides.add_slide(slide_layout)
    
    # 深度复制所有形状
    for shape in source_slide.shapes:
        el = shape.element
        newel = copy.deepcopy(el)
        new_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
    
    # 移动到正确位置（紧跟在原页面后）
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[-1])
    xml_slides.insert(slide_number, slides[-1])
    
    # 保存
    prs.save(output_file)
    print(f"已在第 {slide_number} 页后插入副本，文件已保存: {output_file}")
    return True


def set_pptx_page_texts(pptx_file, output_file, slide_number, new_texts, change_index=None, change_text=None):
    """
    修改指定页的文字内容
    
    Args:
        pptx_file: 原PPTX文件路径
        output_file: 输出PPTX文件路径
        slide_number: 页码（从1开始）
        replacements: 字典，格式 {'旧文字': '新文字'}
    
    Returns:
        bool: 是否成功
    """
    if not os.path.exists(pptx_file):
        print(f"错误：找不到文件 {pptx_file}")
        return False
    
    prs = Presentation(pptx_file)
    
    # 检查页码是否有效
    if slide_number < 1 or slide_number > len(prs.slides):
        print(f"错误：页码 {slide_number} 超出范围（共 {len(prs.slides)} 页）")
        return False
    
    # 获取指定页（索引从0开始）
    slide = prs.slides[slide_number - 1]
    
    # 遍历该页的所有形状
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                if change_index is not None:
                    print(f"{paragraph.runs[change_index].text} → {change_text}")
                    #paragraph.runs[change_index].text = change_text 
                '''
                else:
                    for run in paragraph.runs:
                        # Overwrite text
                        #print(f"{run.text} → ", end="\n")
                        #run.text = new_texts.pop(0) if new_texts else ""
                '''
    
    # 保存
    #prs.save(output_file)
    #print(f"已修改第 {slide_number} 页，文件已保存: {output_file}")
    return True

if __name__ == "__main__":
    # 示例1：读取PPT信息
    filename = "template"
    pptx_file = f"C:\\Users\\eglis\\Desktop\\PPT\\SildePPT\\{filename}.pptx"
    output_file = f"C:\\Users\\eglis\\Desktop\\PPT\\SildePPT\\{filename}_modified.pptx"
    info = read_pptx(output_file)
    
    # 1 时间
    page_to_modify = 1
    date = "04/01/2026"
    old_date = "28/12/2025"
    #update_slide_text(output_file, output_file, page_to_modify, {old_date: date})

    # 2 领会
    '''
20　神能照着运行在我们心里的大力充充足足地成就一切，超过我们所求所想的。 21但愿他在教会中，并在基督耶稣里，得着荣耀，直到世世代代，永永远远。阿们！
    '''
    page_to_modify = 2
    old_name = "徐霞"
    new_name = "周国莲"
    #update_slide_text(output_file, output_file, page_to_modify, {old_name: new_name})
    set_pptx_page_texts(output_file, output_file, page_to_modify, [
        ], change_index=1, change_text=new_name) 

    # 3 敬拜
    page_to_modify = 3
    old_name = "于福芬"
    new_name = "徐霞"
    #update_slide_text(output_file, output_file, page_to_modify, {old_name: new_name})  

    '''
    page_to_delete = [4,4,4,11]
    print(f"删除music页")
    for page in page_to_delete:
        delete_slide(pptx_file, output_file, page)
    '''

    
    '''
    titles = []
    texts = []
    assert len(titles) == len(texts)
    for i in len(titles):
        duplicate_slide(output_file, output_file, page_to_modify)
        page_to_modify += 1
        set_pptx_page_texts(output_file, output_file, page_to_modify, [
            titles[i],
            texts[i]
        ])
    # 示例2：修改单页内容
    #update_slide_text(output_file, output_file, 4, {"主 我敬拜祢": "感谢神"})
    #delete_slide(output_file, output_file, page_to_modify)
    #duplicate_slide(output_file, output_file, page_to_modify)
    '''
    
    '''
    set_pptx_page_texts(output_file, output_file, page_to_modify, [
        #f"{page_to_modify}", 
        "感谢神",
        "感谢神赐喜乐忧愁",
        "感谢神属天平安",
        "感谢神赐明天盼望",
        "要感谢直到永远"])
    '''
    
    # 示例3：批量修改多页
    # slide_replacements = {
    #     1: {'2025': '2026'},
    #     2: {'标题': '新标题'},
    #     3: {'内容': '新内容'}
    # }
    # update_multiple_slides("原文件.pptx", "修改后.pptx", slide_replacements)
    
    # 示例2：修改PPT文字
    # replacements = {
    #     "2025": "2026",
    #     "旧标题": "新标题"
    # }
    # update_pptx_text("原文件.pptx", "修改后.pptx", replacements)
