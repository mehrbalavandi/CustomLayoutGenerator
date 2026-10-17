#!/usr/bin/env python3
"""
🌟 پیدا کردنِ جای یک متن در فایلِ Word (برای گزارشِ ایرادهای داده‌ی کتاب).

شماره‌ی «[ص N]» در پرامپت‌های enrich.py شماره‌ی صفحه‌ی خروجیِ استخراج (اپ) است، نه صفحه‌ی Word.
این ابزار سند را با LibreOffice به PDF تبدیل می‌کند (یک‌بار؛ نتیجه کنارِ سند در پوشه‌ی موقتِ
سیستم کش می‌شود) و برای هر عبارت شماره‌ی صفحه‌ی Word و کمی از متنِ دورش را چاپ می‌کند.
صفحه‌بندیِ LibreOffice معمولاً با Word یکی است، ولی ممکن است یکی-دو صفحه جابه‌جا باشد؛ برای
همین عبارتِ جست‌وجو (که در Word با Ctrl+F پیدا می‌شود) مهم‌تر از شماره است.

    python3 locate.py "Mindset1- Custom- Units 01-08.docx" "ccuntries" "home animal"
"""
import hashlib
import os
import re
import subprocess
import sys
import tempfile
import zipfile


def pdf_for(docx: str) -> str:
    st = os.stat(docx)
    key = hashlib.sha1(f"{os.path.abspath(docx)}|{st.st_mtime}|{st.st_size}".encode()).hexdigest()[:12]
    out_dir = os.path.join(tempfile.gettempdir(), "locate_" + key)
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    if not os.path.exists(pdf):
        os.makedirs(out_dir, exist_ok=True)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=900)
    return pdf


def pdf_pages(pdf: str) -> list[str]:
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
    pages = []
    for i in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), pdf, "-"], capture_output=True, text=True).stdout
        pages.append(re.sub(r"\s+", " ", t))
    return pages


def docx_paragraphs(docx: str) -> list[str]:
    xml = zipfile.ZipFile(docx).read("word/document.xml").decode("utf-8")
    return ["".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)) for p in re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.S)]


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    docx, queries = sys.argv[1], sys.argv[2:]
    paras = docx_paragraphs(docx)
    pages = pdf_pages(pdf_for(docx))
    norm = lambda s: re.sub(r"\s+", " ", s).lower()
    for q in queries:
        nq = norm(q)
        in_docx = [i for i, p in enumerate(paras) if nq in norm(p)]
        on_pages = [i + 1 for i, t in enumerate(pages) if nq in norm(t)]
        print(f"== «{q}»: {len(in_docx)} بار در سند؛ صفحه‌ی Word: {', '.join(map(str, on_pages)) or 'پیدا نشد (شاید در جدول/چند خط)'}")
        for i in in_docx[:5]:
            print(f"   ¶{i}: {paras[i][:160]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
