# -*- coding: utf-8 -*-
"""Trich bai thu hoach mau tu so tay HTML -> .md va .docx"""
import re, html, os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

BASE = r'D:\HocTap\CamTinhDang_2026'
src = open(os.path.join(BASE, '01_SoTay_OnTap_CamTinhDang.html'), encoding='utf-8').read()

start = src.find('<p><b>B\u00c0I THU HO\u1ea0CH</b>')
end = src.find('<h3>3. B\u1ea3n r\u00fat g\u1ecdn 2 trang')
assert start > 0 and end > start, (start, end)
block = src[start:end]

paras = re.findall(r'<p[^>]*>(.*?)</p>', block, flags=re.S)

def clean(t):
    t = re.sub(r'<br\s*/?>', '\n', t)
    t = re.sub(r'<b>(.*?)</b>', r'**\1**', t, flags=re.S)
    t = re.sub(r'<i>(.*?)</i>', r'\1', t, flags=re.S)
    t = re.sub(r'<span[^>]*>', '', t)
    t = t.replace('</span>', '')
    t = re.sub(r'<[^>]+>', '', t)
    t = html.unescape(t)
    return t.strip()

paras = [clean(p) for p in paras]
paras = [p for p in paras if p]
print('paragraphs:', len(paras))

# ---------- Markdown ----------
md = []
md.append('# Bài thu hoạch mẫu — Lớp bồi dưỡng nhận thức về Đảng\n')
md.append('> Học viên: Nguyễn Tuấn Minh — Giáo viên Hóa học, Trường THPT Chi Lăng, tỉnh Gia Lai  ')
md.append('> Bản mẫu để đối chiếu cách viết. Khi làm bài thật phải thay bằng chi tiết của chính mình.\n')
for p in paras:
    md.append(p + '\n')
md_path = os.path.join(BASE, '02_BaiThuHoach_Mau.md')
open(md_path, 'w', encoding='utf-8').write('\n'.join(md))
print('wrote', md_path, os.path.getsize(md_path))

# ---------- DOCX ----------
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin, sec.right_margin = Cm(3), Cm(2)
sec.top_margin, sec.bottom_margin = Cm(2), Cm(2)

normal = doc.styles['Normal']
normal.font.name = 'Times New Roman'
normal.font.size = Pt(13)
normal.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
normal.paragraph_format.line_spacing = 1.5
normal.paragraph_format.space_after = Pt(6)

def add_runs(par, text, base_bold=False, size=None):
    parts = text.split('**')
    for i, part in enumerate(parts):
        if part == '':
            continue
        lines = part.split('\n')
        for j, line in enumerate(lines):
            run = par.add_run(line)
            run.bold = base_bold or (i % 2 == 1)
            if size:
                run.font.size = size
            if j < len(lines) - 1:
                run.add_break()

# Title block: first paragraph = "BAI THU HOACH\nHoc vien: ..."
first = paras[0].split('\n')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run(first[0].strip())
r.bold = True
r.font.size = Pt(15)

for line in first[1:]:
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(line.strip())
    r2.italic = True
    r2.font.size = Pt(12)

for text in paras[1:]:
    is_heading = bool(re.match(r'^\*\*(M\u1edf b\u00e0i|I{1,3}\.|K\u1ebft b\u00e0i)', text))
    p = doc.add_paragraph()
    if is_heading:
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.first_line_indent = Cm(0)
        add_runs(p, text, base_bold=True)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1)
        add_runs(p, text)

docx_path = os.path.join(BASE, '02_BaiThuHoach_Mau.docx')
doc.save(docx_path)
print('wrote', docx_path, os.path.getsize(docx_path))
