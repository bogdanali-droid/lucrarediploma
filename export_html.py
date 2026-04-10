#!/usr/bin/env python3
"""Exportă teza ca HTML complet, self-contained."""
import re, html

BASE = "/home/user/lucrarediploma/teza-master"
CHAPTERS = [
    BASE + "/introducere.md",
    BASE + "/capitolul_1.md",
    BASE + "/capitolul_2.md",
    BASE + "/capitolul_3.md",
    BASE + "/concluzie.md",
    BASE + "/bibliografie.md",
    BASE + "/anexa_1_proiecte_didactice.md",
    BASE + "/anexa_2_instrumente_cercetare.md",
]

def md_to_html(text):
    lines = text.split('\n')
    out = []
    in_table = False
    for line in lines:
        if line.startswith('> ') or line.strip() == '---':
            continue
        if re.match(r'^\|', line):
            if not in_table:
                out.append('<table>')
                in_table = True
            if re.match(r'^\|[-| ]+\|', line):
                continue
            cells = [c.strip() for c in line.strip('|').split('|')]
            out.append('<tr>' + ''.join(f'<td>{inline(c)}</td>' for c in cells) + '</tr>')
            continue
        if in_table:
            out.append('</table>')
            in_table = False
        if line.startswith('# ') and not line.startswith('## '):
            out.append(f'<h1>{html.escape(line[2:])}</h1>')
        elif line.startswith('## ') and not line.startswith('### '):
            out.append(f'<h2>{html.escape(line[3:])}</h2>')
        elif line.startswith('### '):
            out.append(f'<h3>{html.escape(line[4:])}</h3>')
        elif line.startswith('#### '):
            out.append(f'<h4>{html.escape(line[5:])}</h4>')
        elif re.match(r'^(\d+)\.\s', line):
            out.append(f'<p class="biblio">{inline(line)}</p>')
        elif line.startswith('- ') or line.startswith('* '):
            out.append(f'<li>{inline(line[2:])}</li>')
        elif line.startswith('[FOTO'):
            out.append(f'<p class="foto">{html.escape(line.strip())}</p>')
        elif line.strip() == '':
            out.append('<br>')
        else:
            out.append(f'<p>{inline(line.strip())}</p>')
    if in_table:
        out.append('</table>')
    return '\n'.join(out)

def inline(text):
    text = html.escape(text)
    text = re.sub(r'\*\*\*(.*?)\*\*\*', r'<b><i>\1</i></b>', text)
    text = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)
    text = re.sub(r'\*(.*?)\*', r'<i>\1</i>', text)
    return text

body = ""
for i, f in enumerate(CHAPTERS):
    with open(f, encoding='utf-8') as fh:
        content = fh.read()
    body += ('<hr>' if i > 0 else '') + md_to_html(content)

html_out = f"""<!DOCTYPE html>
<html lang="ro">
<head>
<meta charset="UTF-8">
<title>Teza — Anca Aliciuc</title>
<style>
  body {{ font-family: "Times New Roman", serif; font-size: 12pt;
         max-width: 170mm; margin: 25mm auto 25mm auto; line-height: 1.5;
         color: #000; background: #fff; }}
  h1 {{ font-size: 14pt; text-align: center; text-transform: uppercase;
       margin: 24pt 0 12pt 0; }}
  h2 {{ font-size: 13pt; margin: 18pt 0 6pt 0; }}
  h3 {{ font-size: 12pt; font-style: italic; margin: 12pt 0 4pt 0; }}
  h4 {{ font-size: 12pt; margin: 8pt 0 2pt 0; }}
  p  {{ text-align: justify; text-indent: 1.25cm; margin: 0 0 4pt 0; }}
  p.biblio {{ text-indent: -1cm; padding-left: 1cm; }}
  p.foto {{ color: #888; font-style: italic; font-size: 10pt;
            text-align: center; text-indent: 0; border: 1px dashed #ccc;
            padding: 8pt; margin: 8pt 0; }}
  li {{ margin-left: 1cm; margin-bottom: 2pt; }}
  table {{ border-collapse: collapse; width: 100%; margin: 8pt 0; font-size: 10pt; }}
  td {{ border: 1px solid #999; padding: 4pt 6pt; vertical-align: top; }}
  hr {{ margin: 40pt 0; border: none; border-top: 1px solid #ccc; }}
  @media print {{
    body {{ margin: 0; }}
    hr {{ page-break-before: always; border: none; }}
  }}
</style>
</head>
<body>
<h1>STREAM și trans/interdisciplinaritatea în ciclul primar</h1>
<p style="text-align:center;text-indent:0"><b>Anca Aliciuc</b><br>Lucrare de disertație — UPSC Ion Creangă, Chișinău<br>2026</p>
<hr>
{body}
</body>
</html>"""

out = "/home/user/lucrarediploma/teza_anca_aliciuc.html"
with open(out, 'w', encoding='utf-8') as f:
    f.write(html_out)
print(f"Salvat: {out} ({len(html_out)//1024}KB)")
