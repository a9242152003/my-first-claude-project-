#!/usr/bin/env python3
"""Convert the Word document to a styled HTML file for online preview."""

import mammoth

docx_path = '/home/user/my-first-claude-project-/Tintinalli_Ch59-61_中文筆記.docx'
html_out = '/home/user/my-first-claude-project-/Tintinalli_Ch59-61_中文筆記.html'

with open(docx_path, 'rb') as f:
    result = mammoth.convert_to_html(f)
    body = result.value

css = """
<style>
  body {
    font-family: "Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", sans-serif;
    max-width: 900px;
    margin: 2em auto;
    padding: 0 1.5em;
    line-height: 1.75;
    color: #222;
    background: #fafafa;
  }
  h1 { color: #1F4E79; border-bottom: 3px solid #1F4E79; padding-bottom: 0.3em; margin-top: 1.8em; }
  h2 { color: #2E74B5; border-left: 5px solid #2E74B5; padding-left: 0.5em; margin-top: 1.4em; }
  h3 { color: #44546A; margin-top: 1.2em; }
  strong { color: #C00000; }
  ul { padding-left: 1.5em; }
  li { margin: 0.25em 0; }
  p { margin: 0.4em 0; }
  hr { border: none; border-top: 1px dashed #ccc; margin: 2em 0; }
  .header {
    text-align: center;
    padding: 1em;
    background: linear-gradient(90deg, #1F4E79, #2E74B5);
    color: white;
    border-radius: 8px;
    margin-bottom: 2em;
  }
  .header h1 { color: white; border: none; }
  .header p { font-size: 0.95em; opacity: 0.9; }
</style>
"""

html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>Tintinalli Ch 59-61 中文筆記</title>
{css}
</head>
<body>
<div class="header">
  <h1>Tintinalli 急診醫學 第 59–61 章重點筆記</h1>
  <p>主動脈剝離 ‧ 動脈瘤疾病 ‧ 動脈阻塞 | 住院醫師臨床參考</p>
</div>
{body}
</body>
</html>
"""

with open(html_out, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Saved: {html_out}')
