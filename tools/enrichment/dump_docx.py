# -*- coding: utf-8 -*-
"""
ابزارِ کمکیِ غنی‌سازی (فقط برای آماده‌کردنِ متن برای ترجمه؛ اکسترکتور به آن نیازی ندارد).

متنِ هر پاراگرافِ docx را به ترتیبِ سند، با شماره‌ی صفحه (مثلِ اکسترکتور: lastRenderedPageBreak
یا شکستِ صفحه)، سطحِ ساختار (Nav 1 / Nav 2) و علامت‌های {blk}…{/blk} برای جای‌خالی‌ها
(BlkPa / BlkLp: کلِ پاراگراف، BlkCh: یک ران) بیرون می‌دهد. سلول‌های جدول هم به ترتیب.

کلیدِ هر پاراگراف همان normalize() است که ContentEnrichment.cs (اکسترکتور) استفاده می‌کند.

استفاده:
    python3 dump_docx.py "Mindset1- Custom- Units 01-08.docx" out.json [startPage]
"""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def normalize(text: str) -> str:
    """همان قاعده‌ی ContentEnrichment.Normalize در C# — هر تغییری این‌جا، آن‌جا هم."""
    t = text.replace('{blk}', '').replace('{/blk}', '')
    t = (t.replace('‘', "'").replace('’', "'").replace('“', '"')
          .replace('”', '"').replace('–', '-').replace('—', '-')
          .replace(' ', ' ').replace('​', ''))
    t = re.sub(r'\s+', ' ', t).strip().lower()
    return t


def load_styles(z):
    root = ET.fromstring(z.read('word/styles.xml'))
    styles = {}
    for s in root.findall(W + 'style'):
        sid = s.get(W + 'styleId')
        name = s.find(W + 'name')
        based = s.find(W + 'basedOn')
        outline = s.find(f'{W}pPr/{W}outlineLvl')
        styles[sid] = {
            'name': name.get(W + 'val') if name is not None else sid,
            'based': based.get(W + 'val') if based is not None else None,
            'outline': int(outline.get(W + 'val')) if outline is not None else None,
        }
    return styles


def style_chain(styles, sid):
    seen = []
    while sid and sid in styles and sid not in seen:
        seen.append(sid)
        sid = styles[sid]['based']
    return seen


def is_style(styles, sid, target):
    return any(styles[s]['name'] == target or s == target for s in style_chain(styles, sid))


def outline_of(styles, p):
    ppr = p.find(W + 'pPr')
    if ppr is not None:
        o = ppr.find(W + 'outlineLvl')
        if o is not None:
            return int(o.get(W + 'val'))
        ps = ppr.find(W + 'pStyle')
        if ps is not None:
            for s in style_chain(styles, ps.get(W + 'val')):
                if styles[s]['outline'] is not None:
                    return styles[s]['outline']
    return None


def main():
    src, out = sys.argv[1], sys.argv[2]
    page = int(sys.argv[3]) if len(sys.argv) > 3 else 1
    z = zipfile.ZipFile(src)
    styles = load_styles(z)
    body = ET.fromstring(z.read('word/document.xml')).find(W + 'body')
    items = []
    nav = {0: None, 1: None}
    state = {'page': page}

    def para_text(p):
        ppr = p.find(W + 'pPr')
        psid = None
        if ppr is not None and ppr.find(W + 'pStyle') is not None:
            psid = ppr.find(W + 'pStyle').get(W + 'val')
        whole_blank = psid is not None and (is_style(styles, psid, 'BlkPa') or is_style(styles, psid, 'BlkLp'))
        parts = []
        page_break = False
        for r in p.iter(W + 'r'):
            if r.find(W + 'lastRenderedPageBreak') is not None or any(
                    b.get(W + 'type') == 'page' for b in r.findall(W + 'br')):
                page_break = True
            rpr = r.find(W + 'rPr')
            rsid = None
            if rpr is not None and rpr.find(W + 'rStyle') is not None:
                rsid = rpr.find(W + 'rStyle').get(W + 'val')
            caps = rpr is not None and rpr.find(W + 'caps') is not None and rpr.find(W + 'caps').get(W + 'val') not in ('0', 'false')
            t = ''
            for node in r:
                if node.tag == W + 't':
                    t += node.text or ''
                elif node.tag == W + 'tab':
                    t += '\t'
                elif node.tag == W + 'br' and node.get(W + 'type') != 'page':
                    t += '\n'
            if caps:
                t = t.upper()
            if t and rsid and is_style(styles, rsid, 'BlkCh') and not whole_blank:
                t = '{blk}' + t + '{/blk}'
            parts.append(t)
        text = ''.join(parts).replace('{/blk}{blk}', '')
        if whole_blank and text.strip():
            text = '{blk}' + text + '{/blk}'
        return text, page_break

    def walk(el, in_table):
        for child in el:
            if child.tag == W + 'p':
                text, pb = para_text(child)
                if pb:
                    state['page'] += 1
                lvl = outline_of(styles, child)
                if lvl in (0, 1) and text.strip():
                    nav[lvl] = text.strip()
                    if lvl == 0:
                        nav[1] = None
                if re.search(r'[A-Za-z]', text):
                    items.append({
                        'page': state['page'],
                        'unit': nav[0], 'section': nav[1],
                        'inTable': in_table,
                        'text': text,
                        'key': normalize(text),
                    })
            elif child.tag == W + 'tbl':
                for tr in child.findall(W + 'tr'):
                    for tc in tr.findall(W + 'tc'):
                        walk(tc, True)
            elif child.tag in (W + 'sdt',):
                c = child.find(W + 'sdtContent')
                if c is not None:
                    walk(c, in_table)

    walk(body, False)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=1)
    print(len(items), 'paragraphs; last page', state['page'])


if __name__ == '__main__':
    main()
