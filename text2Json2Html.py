import datetime
import os
import tkinter as tk
from tkinter import filedialog
import re
import time
import random
import json
from pathlib import Path

# 定义正则表达式模式
PATTERN_TITLE = re.compile(r'(?:^|\n\s*)\d+?[\.\。\,\、]')  # 匹配题号
PATTERN_NEWLINE = re.compile(r'\n')  # 匹配换行符
PATTERN_ANSWER = re.compile(r'答案[:：]\s*(.*?)(?=\s*解析[:：]|$)',
                            re.DOTALL)  # 匹配答案
PATTERN_ANALYSIS = re.compile(r'解析[:：]\s*(.*)', re.DOTALL)  # 匹配解析

webTitle = ""

# 获取文件路径
root = tk.Tk()
root.withdraw()
file_path = Path(filedialog.askopenfilename(title="选择题库文本文件", filetypes=[('Text Files', '*.txt')]))
if not file_path.exists():
    print("未选择有效文本文件。")
    exit()

html_path = Path(filedialog.askopenfilename(title="选择HTML模板文件", filetypes=[('Html Files', '*.html')]))
if not html_path.exists():
    print("未选择有效HTML模板文件。")
    exit()

save_path = filedialog.asksaveasfilename(title="保存为",defaultextension='.html',filetypes=[('HTML Files', '*.html')])
if not save_path:
    print("未选择保存路径。")
    exit()

head, tail = os.path.split(save_path)
webTitle = tail.split(".")[0]
tail = datetime.date.today().strftime('%Y%m%d') + "_" + tail
save_path = os.path.join(head, tail)
# 将字符串转换为 Path 对象
save_path = Path(save_path)


try:
    # 读取文件内容
    with file_path.open("r", encoding="UTF-8") as f:
        content = f.read().replace('．', '.')  # 替换全角点号为半角点号

    with html_path.open("r", encoding="UTF-8") as f:
        html_content = f.read()

    # 分割问题列表
    titles = PATTERN_TITLE.split(content)

    # 将每个题目分解为题干、题型、选项、答案，解析并存放到result中
    results = []
    for title_text in titles[1:]:  # 跳过第一个空元素
        # 题目
        parts = PATTERN_NEWLINE.split(title_text, 1)
        title = parts[0]
        details = parts[1] if len(parts) > 1 else ""

        # 选项
        options = re.findall(r'[A-J][\.\。\,\、]\s*(.+?)\s+[\n]?', details)

        # # 答案

        match = PATTERN_ANSWER.search(details)
        if match:
            answer = match.group(1).strip()
            # 匹配解析
            analysis_match = PATTERN_ANALYSIS.search(details)
            if analysis_match:
                analysis = analysis_match.group(1).strip()
            else:
                analysis = ''
        else:
            answer = ''
            analysis = ''

        # 格式化成2016-03-20 11:45:39形式
        results.append({
            'id':
            time.strftime("%Y%m%d%H%M", time.localtime()) +
            str(random.randint(100000, 999999)),
            'title':
            title,
            'options':
            options,
            'answer':
            answer,
            'analysis':
            analysis
        })
    
    html_content = html_content.replace("网页模板",webTitle).replace('[{ 替换 }]', json.dumps(results))


    with open(save_path,'w', encoding='utf-8') as f:
        f.write(html_content)

    print("执行完成")
except Exception as e:
    print(f"发生错误：{e}")
