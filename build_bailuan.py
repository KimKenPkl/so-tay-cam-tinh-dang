# -*- coding: utf-8 -*-
"""
Chuan hoa markdown -> HTML (chen vao so tay) va DOCX cho:
  - 03_BaiLuan_CuoiKhoa.md
  - 04_HoTat_KienThuc_5Bai.md
Khong dung Word COM / pandoc / build_de.py.
"""
import re, html, os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

BASE = r'D:\HocTap\CamTinhDang_2026'

# ---------- 1. Chuan hoa heading trong file markdown ----------
HEADING_MAP = {
    '03_BaiLuan_CuoiKhoa.md': {
        '## Mở bài': None,  # da co
        'I. SỰ RA ĐỜI CỘNG SẢN VIỆT NAM': None,
    },
}

def normalize_md(text, filename):
    """Them marker ## / ### cho cac dong tieu de da biet."""
    if filename.startswith('03_'):
        h2 = [
            'I. SỰ RA ĐỜI CỦA ĐẢNG CỘNG SẢN VIỆT NAM — BƯỚC NGOẶT QUYẾT ĐỊNH CỦA CÁCH MẠNG VIỆT NAM',
            'II. NHỮNG THẮNG LỢI TO LỚN DƯỚI SỰ LÃNH ĐẠO CỦA ĐẢNG VÀ CẢM NHẬN CỦA TÔI',
            'III. HỌC TẬP, NGHIÊN CỨU CƯƠNG LĨNH CHÍNH TRỊ CỦA ĐẢNG — Ý NGHĨA ĐỐI VỚI QUẦN CHÚNG ƯU TÚ VÀ NHẬN THỨC VỀ XÃ HỘI ĐANG XÂY DỰNG',
            'KẾT BÀI — PHƯƠNG HƯỚNG PHẤN ĐẤU',
        ]
        h3 = [
            '1. Nội hàm của khái niệm "bước ngoặt quyết định"',
            '2. Vai trò của Nguyễn Ái Quốc — người làm nên bước ngoặt',
            '3. Phân tích lực lượng giai cấp — cái nền làm nên bước ngoặt',
            '1. Hành trình thắng lợi',
            '2. Ý nghĩa của những thắng lợi',
            '3. Cập nhật Đại hội XIV — thắng lợi đang tiếp diễn',
            '4. Cảm nhận và liên hệ bản thân',
            '1. Cương lĩnh là gì',
            '2. Ý nghĩa đối với quần chúng ưu tú',
            '3. Nhận thức về xã hội mà nhân dân ta đang xây dựng',
        ]
    else:
        h2 = []
        h3 = []
    lines = text.split('\n')
    out = []
    for ln in lines:
        s = ln.strip()
        if s in h2:
            out.append('## ' + s)
        elif s in h3:
            out.append('### ' + s)
        else:
            out.append(ln)
    return '\n'.join(out)

# ---------- 2. Markdown -> HTML fragment ----------
def inline(t):
    t = html.escape(t, quote=False)
    t = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', t)
    t = re.sub(r'(?<!\*)\*([^*]+?)\*(?!\*)', r'<i>\1</i>', t)
    return t

def md_to_html(md, skip_h1=True):
    lines = md.split('\n')
    out, i, n = [], 0, len(lines)
    para = []
    def flush():
        if para:
            out.append('<p>' + ' '.join(inline(x) for x in para) + '</p>')
            para.clear()
    while i < n:
        raw = lines[i]
        s = raw.strip()
        if not s:
            flush(); i += 1; continue
        if s == '---':
            flush(); i += 1; continue
        if s.startswith('# ') and skip_h1:
            flush(); i += 1; continue
        if s.startswith('## '):
            flush(); out.append('<h3>' + inline(s[3:]) + '</h3>'); i += 1; continue
        if s.startswith('### '):
            flush(); out.append('<h4>' + inline(s[4:]) + '</h4>'); i += 1; continue
        if s.startswith('|'):
            flush()
            tbl = []
            while i < n and lines[i].strip().startswith('|'):
                tbl.append(lines[i].strip()); i += 1
            rows = []
            for tl in tbl:
                cells = [c.strip() for c in tl.strip('|').split('|')]
                if all(re.fullmatch(r':?-{2,}:?', c) for c in cells):
                    continue
                rows.append(cells)
            if rows:
                out.append('<table class="t">')
                out.append('<thead><tr>' + ''.join('<th>' + inline(c) + '</th>' for c in rows[0]) + '</tr></thead>')
                out.append('<tbody>')
                for r in rows[1:]:
                    out.append('<tr>' + ''.join('<td>' + inline(c) + '</td>' for c in r) + '</tr>')
                out.append('</tbody></table>')
            continue
        if re.match(r'^-\s', s):
            flush()
            items = []
            while i < n and re.match(r'^-\s', lines[i].strip()):
                items.append(re.sub(r'^-\s+', '', lines[i].strip())); i += 1
            out.append('<ul>' + ''.join('<li>' + inline(it) + '</li>' for it in items) + '</ul>')
            continue
        if re.match(r'^\d+\.\s', s):
            flush()
            items = []
            while i < n and re.match(r'^\d+\.\s', lines[i].strip()):
                items.append(re.sub(r'^\d+\.\s+', '', lines[i].strip())); i += 1
            out.append('<ol class="kv">' + ''.join('<li>' + inline(it) + '</li>' for it in items) + '</ol>')
            continue
        para.append(s); i += 1
    flush()
    return '\n  '.join(out)

def build_section(part_id, part_no, title, body_html, extra_class=''):
    return f'''<section class="part" id="{part_id}">
  <div class="part-head">
    <div class="no">PHẦN {part_no}</div>
    <h2>{html.escape(title)}</h2>
  </div>
  <div class="box gold">
    <div class="label gold">Bản hoàn chỉnh — dùng để ôn và đối chiếu</div>
    <p>Nội dung dưới đây là bản đã biên soạn đầy đủ, trích dẫn có nguồn theo <i>Tài liệu bồi dưỡng nhận thức về Đảng</i> (xb lần thứ 24) và Văn kiện Đại hội XIV. Khi làm bài thật, phải viết lại bằng lời và chi tiết của chính mình — người chấm nhận ra ngay bài chép.</p>
  </div>
  {body_html}
</section>
'''

# ---------- 3. Docx ----------
def add_runs(par, text, base_bold=False):
    # XuLy ca **dam** va *nghieng* (khong de lai dau sao trong DOCX)
    tokens = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith('**') and tok.endswith('**'):
            r = par.add_run(tok[2:-2]); r.bold = True
        elif tok.startswith('*') and tok.endswith('*'):
            r = par.add_run(tok[1:-1]); r.italic = True
        else:
            r = par.add_run(tok); r.bold = base_bold

def make_docx(md, path, title):
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

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    r.bold = True; r.font.size = Pt(15)

    lines = md.split('\n')
    for raw in lines:
        s = raw.strip()
        if not s:
            continue
        if s.startswith('# ') or s == '---':
            continue
        if s.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.first_line_indent = Cm(0)
            add_runs(p, s[3:], base_bold=True)
            continue
        if s.startswith('### '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.first_line_indent = Cm(0)
            add_runs(p, s[4:], base_bold=True)
            continue
        if s.startswith('|'):
            continue  # bang khong dua vao docx de don gian hoa
        p = doc.add_paragraph()
        if re.match(r'^-\s|^\d+\.\s', s):
            p.paragraph_format.first_line_indent = Cm(0)
            add_runs(p, re.sub(r'^-\s+|^\d+\.\s+', '• ', s))
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1)
            add_runs(p, s)
    doc.save(path)
    print('doc:', path, os.path.getsize(path))

# ---------- 4. Main ----------
def main():
    specs = [
        ('03_BaiLuan_CuoiKhoa.md', 'k', 'K', 'Bài luận cuối khóa — Lớp bồi dưỡng nhận thức về Đảng', '03_BaiLuan_CuoiKhoa.docx', 'BÀI LUẬN CUỐI KHÓA — LỚP BỒI DƯỠNG NHẬN THỨC VỀ ĐẢNG'),
        ('04_HoTat_KienThuc_5Bai.md', 'l', 'L', 'Hệ thống hóa kiến thức 5 bài theo nội dung Ban Tuyên giáo', '04_HoTat_KienThuc_5Bai.docx', 'HỆ THỐNG HÓA KIẾN THỨC 5 BÀI — LỚP BỒI DƯỠNG NHẬN THỨC VỀ ĐẢNG'),
    ]
    sections, toc_lines, nav_links = [], [], []
    for mdname, pid, pno, title, docxname, docxtitle in specs:
        mdpath = os.path.join(BASE, mdname)
        raw = open(mdpath, encoding='utf-8').read()
        norm = normalize_md(raw, mdname)
        if norm != raw:
            open(mdpath, 'w', encoding='utf-8').write(norm)
            print('normalized', mdname)
        sections.append(build_section(pid, pno, title, md_to_html(norm)))
        toc_lines.append(f'    <a href="#{pid}"><span>{pno}</span>{html.escape(title)}</a>')
        nav_links.append(f'<a href="#{pid}">{pno}</a>')
        make_docx(norm, os.path.join(BASE, docxname), docxtitle)

    block = '\n<!-- ================= PART K/L ================= -->\n' + '\n'.join(sections)
    toc_block = '\n'.join(toc_lines)
    nav_block = ''.join(nav_links)

    for fname in ('index.html', '01_SoTay_OnTap_CamTinhDang.html'):
        fpath = os.path.join(BASE, fname)
        src = open(fpath, encoding='utf-8').read()
        if f'id="k"' not in src:
            # chen section truoc P1
            marker = '<!-- ================= PART P1 ================= -->'
            assert marker in src, fname
            src = src.replace(marker, block + '\n' + marker, 1)
            # chen nav sau I
            src = src.replace('<a href="#i">I</a>', '<a href="#i">I</a>' + nav_block, 1)
            # chen toc sau dong I
            i_line = '    <a href="#i"><span>I</span>Lịch sử, nhân vật, bối cảnh thế giới và thời sự 11/2026</a>'
            assert i_line in src, fname
            src = src.replace(i_line, i_line + '\n' + toc_block, 1)
            open(fpath, 'w', encoding='utf-8').write(src)
            print('updated', fname, os.path.getsize(fpath))
        else:
            print('skip (already updated)', fname)

if __name__ == '__main__':
    main()
