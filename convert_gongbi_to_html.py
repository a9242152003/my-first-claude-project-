#!/usr/bin/env python3
"""Convert the gongbi docx to Google-Docs-paste-friendly HTML."""

import mammoth

style_map = """
p[style-name='Heading 1'] => h2.lv1:fresh
p[style-name='Heading 2'] => h3.lv2:fresh
p[style-name='Heading 3'] => h4.lv3:fresh
p[style-name='Heading 4'] => h5.lv4:fresh
p[style-name='Heading 5'] => h6.lv5:fresh
p[style-name='Heading 6'] => p.lv6:fresh
p[style-name='內文1'] => p.t1:fresh
p[style-name='內文2'] => p.t2:fresh
p[style-name='內文3'] => p.t3:fresh
p[style-name='內文4'] => p.t4:fresh
p[style-name='內文5'] => p.t5:fresh
p[style-name='內文6'] => p.t6:fresh
"""

src = '/home/user/my-first-claude-project-/Tintinalli_Ch56-61_共筆.docx'
out_html = '/home/user/my-first-claude-project-/Tintinalli_Ch56-61_共筆_GoogleDocs.html'

with open(src, 'rb') as f:
    result = mammoth.convert_to_html(f, style_map=style_map)
    body = result.value

# CSS optimized for Google Docs paste — uses inline-style-ish formatting
css = """
<style>
  body {
    font-family: "Microsoft JhengHei", "PingFang TC", "Noto Sans TC", "Arial Unicode MS", sans-serif;
    max-width: 900px;
    margin: 1em auto;
    padding: 1em;
    line-height: 1.7;
    color: #000;
  }
  h1 { font-size: 22pt; font-weight: bold; color: #1F4E79;
       border-bottom: 2px solid #1F4E79; padding-bottom: 4px; margin-top: 1em; }
  h2.lv1 { font-size: 16pt; font-weight: bold; color: #1F4E79; margin-top: 1em; }
  h3.lv2 { font-size: 14pt; font-weight: bold; color: #2E74B5; margin-left: 1em; }
  h4.lv3 { font-size: 13pt; font-weight: bold; color: #44546A; margin-left: 2em; }
  h5.lv4 { font-size: 12pt; font-weight: bold; margin-left: 3em; }
  h6.lv5 { font-size: 12pt; font-weight: bold; margin-left: 4em; }
  p.lv6  { font-size: 12pt; font-weight: bold; margin-left: 5em; }
  p.t1 { margin-left: 1em; }
  p.t2 { margin-left: 2em; }
  p.t3 { margin-left: 3em; }
  p.t4 { margin-left: 4em; }
  p.t5 { margin-left: 5em; }
  p.t6 { margin-left: 6em; }
  strong { font-weight: bold; color: #C00000; }
  p { margin: 0.3em 0; }
  .topic { background: #D9D9D9; padding: 6px 10px; font-size: 16pt;
           font-weight: bold; margin-top: 1.5em; }
  .header { text-align:center; margin-bottom: 1.5em; }
  .header h1 { border: none; }
  .instructions {
    background: #FFF3CD; border: 1px solid #FFC107;
    padding: 12px; border-radius: 6px; margin-bottom: 1.5em;
    font-size: 11pt; color: #555;
  }
</style>
"""

# Wrap with full HTML page
html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>Tintinalli Ch 56-61 共筆 (Google Docs 可貼上版)</title>
{css}
</head>
<body>

<div class="instructions">
  <strong>📋 貼到 Google Docs 步驟:</strong><br>
  1. 在這個頁面按 <strong>Ctrl+A</strong>(Mac: Cmd+A)全選<br>
  2. 按 <strong>Ctrl+C</strong>(Mac: Cmd+C)複製<br>
  3. 開啟空白 Google 文件 → 按 <strong>Ctrl+V</strong>(Mac: Cmd+V)貼上<br>
  4. 標題、粗體、縮排格式皆會保留
</div>

<div class="header">
  <h1>Tintinalli 急診醫學  第 56–61 章共筆</h1>
  <p><strong>Cardiovascular Disease</strong> | 住院醫師整理 | VTE/PE ‧ Hypertension ‧ Pulmonary Hypertension ‧ Aortic Dissection ‧ Aneurysmal Disease ‧ Arterial Occlusion</p>
</div>

{body}

</body>
</html>
"""

with open(out_html, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Saved: {out_html}')
print(f'Messages: {len(result.messages)} warnings')
