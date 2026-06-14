#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""违禁字符扫描脚本：检查 report.json 中是否包含禁止的 markdown 格式、AI 标签、emoji 等。"""
import json
import re
import sys

# 加载报告
with open('.law-research/output/report.json', 'r', encoding='utf-8') as f:
    report = json.load(f)

# 收集所有文本内容
def collect_text(obj):
    """递归收集所有文本内容。"""
    texts = []
    if isinstance(obj, str):
        texts.append(obj)
    elif isinstance(obj, dict):
        for v in obj.values():
            texts.extend(collect_text(v))
    elif isinstance(obj, list):
        for item in obj:
            texts.extend(collect_text(item))
    return texts

all_text = '\n'.join(collect_text(report))

errors = []

# 1. Markdown 加粗
if re.search(r'\*\*[^*]+\*\*', all_text):
    errors.append("包含 Markdown 加粗 **文字**")

# 2. Markdown 标题
if re.search(r'(?m)^#{1,6}\s', all_text):
    errors.append("包含 Markdown 标题 #")

# 3. Markdown 斜体
if re.search(r'(?<!\*)\*(?!\*)[^*]+\*(?!\*)', all_text):
    errors.append("包含 Markdown 斜体 *文字*")

# 4. Markdown 无序列表
if re.search(r'(?m)^[-*]\s', all_text):
    errors.append("包含 Markdown 无序列表 - / *")

# 5. Markdown 有序列表
if re.search(r'(?m)^\d+\.\s', all_text):
    errors.append("包含 Markdown 有序列表 1.")

# 6. Pipe 字符伪表格
# 注意：合法的 URL 中可能包含 pipe，需要排除 URL 上下文
lines_with_pipe = [l for l in all_text.split('\n') if '|' in l and 'http' not in l]
if lines_with_pipe:
    errors.append(f"包含 Pipe 字符伪表格：{lines_with_pipe[0][:80]}")

# 7. 装饰性分割线
if re.search(r'(?m)^[-=*_]{3,}$', all_text):
    errors.append("包含装饰性分割线 ---/===/***/___")

# 8. 行内代码
if re.search(r'`[^`]+`', all_text):
    errors.append("包含行内代码标记")

# 9. Blockquote
if re.search(r'(?m)^>\s', all_text):
    errors.append("包含 Blockquote >")

# 10. Markdown 链接语法
if re.search(r'\[.*?\]\(.*?\)', all_text):
    errors.append("包含 Markdown 链接语法 [文字](链接)")

# 11. 删除线
if re.search(r'~~.*?~~', all_text):
    errors.append("包含删除线 ~~")

# 12. AI 风格标签
ai_tags = ['【风险评估】', '【合规建议】', '【注意事项】', '【重要提示】', '【分析结论】']
for tag in ai_tags:
    if tag in all_text:
        errors.append(f"包含 AI 风格标签: {tag}")

# 13. Emoji
emoji_pattern = re.compile(
    r'[\U0001F300-\U0001F9FF'  # 杂项符号和象形文字
    r'\U0001FA00-\U0001FA6F'  # 棋牌符号
    r'\U0001FA70-\U0001FAFF'  # 扩展-A
    r'\U00002600-\U000027BF'  # 杂项符号
    r'\U0001F600-\U0001F64F'  # 表情符号
    r'\U0001F680-\U0001F6FF'  # 交通和地图符号
    r'\U0001F900-\U0001F9FF'  # 补充符号和象形文字
    r']',
    re.UNICODE
)
emoji_matches = emoji_pattern.findall(all_text)
if emoji_matches:
    errors.append(f"包含 emoji: {emoji_matches[:5]}")

# 14. ASCII 装饰线
if re.search(r'[-=#*_]{5,}', all_text):
    errors.append("包含 ASCII 装饰线")

# 15. 过度强调标点
if re.search(r'!{3,}', all_text):
    errors.append("包含三个以上感叹号")
if re.search(r'\?{3,}', all_text):
    errors.append("包含三个以上问号")

# 16. AI 风格装饰引导语
ai_guides = ['> ⚠️', '> 📋', '> ❌', '> ✅']
for guide in ai_guides:
    if guide in all_text:
        errors.append(f"包含 AI 风格装饰引导语: {guide}")

if errors:
    print("❌ 违禁字符扫描不通过！发现以下问题：")
    for e in errors:
        print(f"  - {e}")
    sys.exit(1)
else:
    print("✅ 违禁字符扫描通过！所有内容符合规范。")
    sys.exit(0)
