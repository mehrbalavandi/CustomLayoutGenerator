#!/usr/bin/env python3
"""
🌟 مقایسه‌ی متنِ «Audio Scripts.docx» با «Audio Scripts_markers.docx» (بدونِ مارکرها).

اکسترکتور و enrich.py متنِ صوتی را از فایلِ _markers می‌خوانند، پس هر اصلاحی که فقط در فایلِ
اصلیِ Audio Scripts انجام شود به اپ نمی‌رسد. این ابزار برای هر ترک، کلمه‌هایی را که بینِ دو
فایل فرق دارند با کمی از متنِ دورشان چاپ می‌کند، و مارکرهای ناقص را هم گزارش می‌کند (کلمه‌ی
بی‌مارکر، یا مارکرِ اضافه بعد از حذفِ کلمه).

    python3 compare_scripts.py "Mindset 01"
    python3 compare_scripts.py "Mindset1- Audio Scripts.docx" "Mindset1- Audio Scripts_markers.docx"
"""
import difflib
import glob
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
# همان مارکرهایی که CustomLayoutGenerator و AudioMarkerReview می‌شناسند
MARKER = re.compile(r'\[(?:(?:\d{1,2}):)?\d{1,2}:\d{2}(?:\.\d+)?\]|\[\d+\]')


def tracks(path):
    """نامِ ترک (ردیفِ اول) ← متنِ ردیفِ دوم، برای هر جدولِ اسکریپت (مثلِ AudioMarkerReview)."""
    root = ET.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    out = {}
    for tbl in root.iter(W + 'tbl'):
        rows = tbl.findall(W + 'tr')
        if len(rows) < 2:
            continue
        name = ''.join(t.text or '' for t in rows[0].iter(W + 't')).strip()
        if not name:
            continue
        paras = []
        for p in rows[1].iter(W + 'p'):
            s = ''
            for el in p.iter():
                if el.tag == W + 't':
                    s += el.text or ''
                elif el.tag in (W + 'br', W + 'tab'):
                    s += ' '
            paras.append(s)
        out[name] = '\n'.join(paras)
    return out


def marker_problem(raw):
    """None اگر «یک مارکر پیش از کلمه‌ی اول و یکی پس از هر کلمه» برقرار است، وگرنه توضیح."""
    marks, clean, last = [], '', 0
    for m in MARKER.finditer(raw):
        clean += raw[last:m.start()]
        marks.append(len(clean))
        last = m.end()
    clean += raw[last:]
    if len(marks) < 2:
        return None  # هنوز مارک نخورده
    toks = list(re.finditer(r'\S+', clean))
    ends = set(marks)
    bare = [t.group(0) for t in toks if t.end() not in ends]
    if bare:
        return 'کلمه‌ی بی‌مارکر: ' + '، '.join(repr(b) for b in bare[:5])
    if len(marks) != len(toks) + 1:
        return f'{len(marks) - len(toks) - 1} مارکرِ اضافه (احتمالاً کلمه‌ای حذف شده و مارکرش مانده)'
    return None


def main(argv):
    if len(argv) == 1 and os.path.isdir(argv[0]):
        src = [f for f in glob.glob(os.path.join(argv[0], '*Audio Scripts.docx'))]
        mk = glob.glob(os.path.join(argv[0], '*Audio Scripts_markers.docx'))
        if len(src) != 1 or len(mk) != 1:
            sys.exit('در این پوشه دقیقاً یک «*Audio Scripts.docx» و یک «*Audio Scripts_markers.docx» لازم است.')
        argv = [src[0], mk[0]]
    if len(argv) != 2:
        sys.exit(__doc__)
    a_all, b_all = tracks(argv[0]), tracks(argv[1])
    diffs = 0
    for name in sorted(set(a_all) | set(b_all)):
        if name not in b_all or name not in a_all:
            print(f'== {name}: فقط در {"فایلِ اصلی" if name in a_all else "فایلِ _markers"}')
            continue
        a = MARKER.sub('', a_all[name]).split()
        b = MARKER.sub('', b_all[name]).split()
        lines = []
        for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
            if op == 'equal':
                continue
            lines.append(f'   اصلی: «{" ".join(a[max(0, i1 - 3):i2 + 3])}»\n'
                         f'   _markers: «{" ".join(b[max(0, j1 - 3):j2 + 3])}»')
        problem = marker_problem(b_all[name])
        if lines or problem:
            print(f'== {name}')
            if problem:
                print('   ⚠️ ' + problem)
            print('\n'.join(lines))
            diffs += len(lines)
    print(f'\nمجموع: {diffs} اختلافِ متنی.')


if __name__ == '__main__':
    main(sys.argv[1:])
