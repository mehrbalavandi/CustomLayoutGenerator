# -*- coding: utf-8 -*-
"""
ابزارِ غنی‌سازی، مستقل از اینکه کدام هوش مصنوعی ترجمه می‌کند (Claude، Gemini، ChatGPT، …).

فقط کتابخانه‌ی استانداردِ پایتون ۳ لازم است. دستورها (از پوشه‌ی tools/enrichment):

  status   آمارِ ترجمه‌ی یک کتاب (به تفکیکِ Unit و اسکریپت‌های صوتی)
      python enrich.py status  "…/assets/data/Mindset 01"

  prepare  پاراگراف‌های بدونِ ترجمه را واحد به واحد به «پرامپت»های آماده تبدیل می‌کند
      python enrich.py prepare "…/assets/data/Mindset 01" [--unit "UNIT/02"] [--audio] [--max-items 80] [--from-pending]
      ← {کتاب}/enrichment/batches/batch_NNN.prompt.md  (همین را به هوش مصنوعی بدهید)
        {کتاب}/enrichment/batches/batch_NNN.ids.json   (نگاشتِ شناسه‌ها؛ دست نزنید)

  merge    پاسخِ JSONِ هوش مصنوعی را بررسی و به translations.json و lexicon.json اضافه می‌کند
      python enrich.py merge   "…/assets/data/Mindset 01" batch_003 پاسخ.json [--dry-run]

قواعدِ ترجمه و انتخابِ کلمات از ENRICHMENT-GUIDE.md (بخش‌های ۲ و ۳) خوانده و داخلِ هر
پرامپت گذاشته می‌شود؛ پس فقط همان یک فایل را به‌روز نگه دارید.
"""
import argparse
import glob
import json
import os
import re
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import dump_docx  # noqa: E402

LEX_REQUIRED = ['headword', 'forms', 'pos', 'ipa', 'pronounceFa', 'pronounceAr', 'translationFa',
                'translationAr', 'explanationFa', 'explanationAr', 'cefrLevel']
NON_WESTERN_DIGITS = str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩', '01234567890123456789')

# شماره‌ی جای خالی: عددِ یک یا دو رقمیِ تنها (شاید داخلِ پرانتز) درست پیش از {blk}، مثلِ
# «They are a 2 {blk}close{/blk} happy family.» — شماره‌ی سؤال است، نه بخشی از جمله
# (ENRICHMENT-GUIDE، قاعده‌ی ۲-۵).
BLANK_NUMBER = re.compile(r'(?:(?<=\s)|^)\(?(\d{1,2})\)?\s*(?=\{blk\})')


# جای خالیِ بی‌استایل: فقط خط‌تیره/زیرخط/[ ... ]/•••، بدونِ {blk} (ENRICHMENT-GUIDE، قاعده‌ی ۲-۵).
# پاسخِ هر جای خالی (G-8)
BLANK_ANSWER = re.compile(r'\{blk\}(.*?)\{/blk\}', re.S)
PLAIN_BLANK = re.compile(r'-{2,}|_{2,}|(?:_ ){2,}_|\[\s*(?:\.{3}|…)\s*\]|\.{4,}|•{3,}')


def blank_numbers(text):
    return BLANK_NUMBER.findall(text)


# ───────────────────────── فایل‌ها ─────────────────────────

def load_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, encoding='utf-8-sig') as f:
        return json.load(f)


def save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write('\n')
    os.replace(tmp, path)


def find_lexicon(book_dir):
    d = os.path.abspath(book_dir)
    for _ in range(4):
        p = os.path.join(d, '_lexicon', 'lexicon.json')
        if os.path.exists(p):
            return p
        d = os.path.dirname(d)
    # پیش‌فرض: کنارِ پوشه‌ی کتاب‌ها
    return os.path.join(os.path.dirname(os.path.abspath(book_dir)), '_lexicon', 'lexicon.json')


def tm_path(book_dir):
    return os.path.join(book_dir, 'enrichment', 'translations.json')


def all_translated_keys(book_dir):
    """کلیدهای ترجمه‌شده در این کتاب و کتاب‌های هم‌سطح (G-2)."""
    keys = set()
    parent = os.path.dirname(os.path.abspath(book_dir))
    for f in glob.glob(os.path.join(parent, '*', 'enrichment', 'translations.json')):
        for k in load_json(f, {}).get('entries', {}):
            keys.add(dump_docx.normalize(k))
    return keys


def book_items(book_dir, audio):
    pattern = '*Audio Scripts_markers.docx' if audio else '*Custom*.docx'
    files = [f for f in glob.glob(os.path.join(book_dir, pattern)) if not os.path.basename(f).startswith('~$')]
    if not files:
        sys.exit(f'فایلِ {pattern} در «{book_dir}» پیدا نشد.')
    out = os.path.join(book_dir, 'enrichment', 'batches', '_dump_audio.json' if audio else '_dump.json')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    argv = sys.argv
    sys.argv = ['dump_docx.py'] + (['--audio'] if audio else []) + [files[0], out]
    try:
        dump_docx.main()
    finally:
        sys.argv = argv
    return load_json(out, [])


def guide_rules():
    text = open(os.path.join(HERE, 'ENRICHMENT-GUIDE.md'), encoding='utf-8').read()
    start = text.find('## ۲.')
    end = text.find('## ۴.')
    return text[start:end].strip() if start >= 0 else text


# ───────────────────────── status ─────────────────────────

def audio_units(book_dir):
    """نامِ فایلِ صوتی ← Unitِ کتاب (از لینک‌های صوتیِ docxِ کتاب)."""
    book_items(book_dir, False)
    return dict(dump_docx.AUDIO_UNITS)


def cmd_status(a):
    done = all_translated_keys(a.book)
    units = {}
    for audio in (False, True):
        if audio:
            units = audio_units(a.book)
        try:
            items = book_items(a.book, audio)
        except SystemExit:
            continue
        groups = {}
        for x in items:
            g = x['section'] if audio else (x['unit'] or '(پیش از اولین Unit)')
            t = groups.setdefault(g, [set(), set()])
            t[0].add(x['key'])
            if x['key'] in done:
                t[1].add(x['key'])
        print('\n=== ' + ('اسکریپت‌های صوتی' if audio else 'متنِ کتاب') + ' ===')
        tot = fin = 0
        for g, (k, d) in groups.items():
            tot += len(k)
            fin += len(d)
            mark = '✓' if len(d) == len(k) else ' '
            unit = f'  ({(units.get(g) or "?")[:20]})' if audio else ''
            print(f'{mark} {g[:48]:48s} {len(d):5d} / {len(k):5d}{unit}')
        print(f'  مجموع (متن‌های یکتا): {fin} / {tot}')


# ───────────────────────── prepare ─────────────────────────

PROMPT_HEAD = """# وظیفه: ترجمه‌ی فارسی و عربی + انتخابِ کلمات و اصطلاحاتِ تعاملی (کتابِ آیلتس)

شما متخصصِ آموزشِ زبانِ انگلیسی (آیلتس) و مترجمِ حرفه‌ای فارسی و عربی هستید. پاراگراف‌های
بخشِ «برای ترجمه» را طبقِ قواعدِ زیر ترجمه کنید و کلمات/اصطلاحاتِ تعاملیِ هر پاراگراف را
انتخاب کنید. بخشِ «متنِ کاملِ واحد» فقط برای فهمِ بافت و پیدا کردنِ پاسخِ جای‌خالی‌هاست
(پاسخ‌های پنهان بینِ {{blk}} و {{/blk}} آمده‌اند)؛ آن را ترجمه نکنید.

{rules}

## قالبِ خروجی (فقط و فقط همین JSON، بدونِ هیچ متنِ دیگر)

```json
{{
  "translations": [
    {{ "id": "p1", "fa": "…", "ar": "…", "terms": ["شناسه‌ی مدخل", "…"] }},
    {{ "id": "p2", "fa": "…", "ar": "…", "terms": [], "blanks": [{{ "fa": "…", "ar": "…" }}] }}
  ],
  "newLexicon": {{
    "شناسه‌ی مدخلِ تازه": {{
      "headword": "…", "forms": ["…"], "pos": "…", "ipa": "/…/",
      "pronounceFa": "…", "pronounceAr": "…",
      "translationFa": "…", "translationAr": "…",
      "explanationFa": "…", "explanationAr": "…",
      "cefrLevel": "B2"
    }}
  }}
}}
```

- برای **همه‌ی** شناسه‌های بخشِ «برای ترجمه» یک آیتم در translations بدهید.
- `plainBlank: true` (اگر در آیتمی آمده): جای خالی بدونِ {{blk}}، فقط با خط‌تیره یا زیرخط. طبقِ
  قاعده‌ی «جای خالیِ بی‌استایل»، اگر پاسخش در بافت هست با پاسخ ترجمه کنید، وگرنه `---` بماند.
- `blankAnswers` (اگر در آیتمی آمده): پاسخِ جای‌خالی‌های همان پاراگراف، به ترتیب. برای هر
  کدام یک آیتم در `blanks` بدهید (همان تعداد، همان ترتیب): ترجمه‌ی **خودِ پاسخ**، کوتاه و با
  همان نقشِ دستوری در جمله (مثلاً `are picked` ← «چیده می‌شوند»، `niece` ← «دختر برادر/خواهر») —
  نه کلِ جمله. اگر پاسخ فقط حرف/شماره‌ی گزینه، عدد، نام، یا کلمه‌ای است که تنها ترجمه‌ی
  معناداری ندارد (of، the)، `{{"fa": "", "ar": ""}}`. اپ در پنجره‌ی پاسخ فقط همین را نشان
  می‌دهد تا پاسخ‌های دیگرِ جمله لو نرود.
- `blankNumbers` (اگر در آیتمی آمده): عددهایی که درست پیش از یک جای خالی‌اند و به احتمالِ
  زیاد فقط شماره‌ی آن جای خالی در تمرین‌اند. طبقِ قاعده‌ی «شماره‌ی جای خالی» در ترجمه
  نمی‌آیند — مگر بافت نشان دهد واقعاً جزءِ جمله‌اند.
- `terms`: شناسه‌ی مدخل‌هایی از دیکشنری که در **همین** پاراگراف تعاملی شوند. اگر مدخل در
  «دیکشنریِ موجود» (پایین) هست، همان شناسه را بنویسید و آن را در newLexicon تکرار نکنید.
  اگر نیست، شناسه‌ی تازه (معمولاً خودِ headword با حروفِ کوچک) بسازید و مدخلِ کاملش را در
  newLexicon بیاورید.
- هر شکلِ `forms` باید عیناً (بی‌توجه به بزرگی/کوچکیِ حروف) در متنِ پاراگرافی آمده باشد که
  آن مدخل را در terms دارد.

## دیکشنریِ موجود (شناسه ← معنا)

{lexicon}

## متنِ کاملِ واحد (فقط برای بافت؛ ترجمه نکنید)

{context}

## برای ترجمه

```json
{items}
```
"""


def cmd_prepare(a):
    items = book_items(a.book, a.audio)
    done = all_translated_keys(a.book)
    lex = load_json(find_lexicon(a.book), {}).get('entries', {})

    # گروه‌بندی: Unit برای متنِ کتاب، فایلِ صوتی برای اسکریپت‌ها
    groups = []
    for x in items:
        g = x['section'] if a.audio else (x['unit'] or '(پیش از اولین Unit)')
        if not groups or groups[-1][0] != g:
            groups.append((g, []))
        groups[-1][1].append(x)
    if a.unit:
        if a.audio:
            # اسکریپت‌های صوتیِ یک Unit: فایل‌هایی که لینکشان در آن Unitِ کتاب است.
            units = audio_units(a.book)
            groups = [g for g in groups if a.unit.lower() in (units.get(g[0]) or '').lower()]
        else:
            groups = [g for g in groups if a.unit.lower() in (g[0] or '').lower()]
    if a.audio:
        # اسکریپت‌ها کوتاه‌اند: چند فایلِ صوتیِ پشتِ‌هم در یک دسته (تا سقفِ max-items)،
        # هرکدام با بافتِ کاملِ خودش.
        packed = []
        for name, gitems in groups:
            if packed and len(packed[-1][1]) + len(gitems) <= a.max_items:
                packed[-1] = (packed[-1][0] + '، ' + name, packed[-1][1] + gitems)
            else:
                packed.append((name, list(gitems)))
        groups = packed

    out_dir = os.path.join(a.book, 'enrichment', 'batches')
    done_dir = os.path.join(out_dir, 'done')
    os.makedirs(done_dir, exist_ok=True)
    # شماره‌ی دسته‌ها پشتِ‌هم می‌ماند (دسته‌های ادغام‌شده هم حساب می‌شوند) و متنی که از قبل
    # در یک دسته‌ی ادغام‌نشده هست، دوباره در دسته‌ی تازه نمی‌آید.
    existing = [int(m) for d in (out_dir, done_dir) for f in os.listdir(d)
                for m in re.findall(r'batch_(\d+)\.ids\.json', f)]
    n = max(existing, default=0)
    for f in glob.glob(os.path.join(out_dir, 'batch_*.ids.json')):
        for v in load_json(f, {}).get('items', {}).values():
            done.add(v['key'])
    lex_lines = '\n'.join(f'- `{k}`: {v.get("headword", k)} — {v.get("translationFa", "")}' for k, v in sorted(lex.items())) or '(خالی)'
    # --from-pending: فهرستِ «چه چیزی ترجمه شود» از pending.jsonِ خودِ اکسترکتور (کلیدهای
    # دقیقِ اکسترکتور، مثلاً برای پاراگراف‌های ادغام‌شده‌ی جای‌خالی)، بافت همچنان از docx.
    pending_by_group = None
    if a.from_pending:
        pend = load_json(os.path.join(a.book, 'enrichment', 'pending.json'), {}).get('items', [])
        pend = [x for x in pend if (x.get('where') == 'audio') == a.audio]
        pending_by_group = {}
        for x in pend:
            g = x['section'] if a.audio else (x.get('unit') or '(پیش از اولین Unit)')
            pending_by_group.setdefault(g, []).append(x)
        known = set()
        for name, _ in groups:
            known.update(name.split('، '))
        for g, xs in pending_by_group.items():
            if g not in known:   # گروهی که در docx پیدا نشد: بدونِ بافتِ جدا
                groups.append((g, xs))

    made = 0
    for name, gitems in groups:
        todo, seen = [], set()
        source = gitems
        if pending_by_group is not None:
            source = [x for g in name.split('، ') for x in pending_by_group.get(g, [])]
        for x in source:
            key = dump_docx.normalize(x['text'])
            if key in done or key in seen:
                continue
            seen.add(key)
            x = dict(x, key=key)
            todo.append(x)
        if not todo:
            continue
        lines, last = [], None
        for x in gitems:
            if a.audio and x['section'] != last:
                last = x['section']
                lines.append(f'\n#### {last}')
            lines.append(x['text'] if a.audio else f'[ص {x["page"]}] {x["text"]}')
        context = '\n'.join(lines)
        for start in range(0, len(todo), a.max_items):
            part = todo[start:start + a.max_items]
            n += 1
            batch = f'batch_{n:03d}'
            ids = {f'p{i + 1}': {'key': x['key'], 'src': x['text'], 'page': x['page'],
                                 'unit': x['unit'], 'section': x['section']} for i, x in enumerate(part)}
            payload = []
            for i, v in ids.items():
                item = {'id': i, 'text': v['src']}
                nums = blank_numbers(v['src'])
                if nums:
                    item['blankNumbers'] = nums
                if PLAIN_BLANK.search(v['src']):
                    item['plainBlank'] = True
                answers = BLANK_ANSWER.findall(v['src'])
                if answers:
                    item['blankAnswers'] = answers
                payload.append(item)
            prompt = PROMPT_HEAD.format(rules=guide_rules(), lexicon=lex_lines,
                                        context=(f'### {name}\n\n' if not a.audio else '') + context,
                                        items=json.dumps(payload, ensure_ascii=False, indent=1))
            with open(os.path.join(out_dir, batch + '.prompt.md'), 'w', encoding='utf-8') as f:
                f.write(prompt)
            save_json(os.path.join(out_dir, batch + '.ids.json'),
                      {'book': os.path.basename(os.path.abspath(a.book)), 'group': name,
                       'audio': a.audio, 'created': datetime.now().strftime('%Y-%m-%d %H:%M'), 'items': ids})
            made += 1
            print(f'{batch}: {name} — {len(part)} پاراگراف')
    print(f'{made} پرامپت در {out_dir}' if made else 'چیزی برای ترجمه نمانده است.')


# ───────────────────────── merge ─────────────────────────

def extract_json(text):
    """پاسخِ هوش مصنوعی ممکن است داخلِ ```json … ``` یا با متنِ اضافه باشد."""
    m = re.search(r'```(?:json)?\s*(\{.*\})\s*```', text, re.S)
    if m:
        text = m.group(1)
    start, end = text.find('{'), text.rfind('}')
    return json.loads(text[start:end + 1])


def clean_nums(src):
    """متن بدونِ شماره‌های جای خالی — تا اگر همان عدد جای دیگری از جمله هم هست، هشدار ندهیم."""
    return BLANK_NUMBER.sub('', src)


def form_in(text, form):
    return re.search(r'(?<![A-Za-z0-9])' + re.escape(form) + r'(?![A-Za-z0-9])', text, re.I) is not None


def cmd_merge(a):
    batch = a.batch.replace('.ids.json', '').replace('.prompt.md', '')
    ids_path = os.path.join(a.book, 'enrichment', 'batches', batch + '.ids.json')
    meta = load_json(ids_path, None)
    if meta is None:
        sys.exit(f'{ids_path} پیدا نشد.')
    resp = extract_json(open(a.response, encoding='utf-8-sig').read())
    lex_path = find_lexicon(a.book)
    lex_root = load_json(lex_path, {'schema': 1, 'entries': {}})
    lex = lex_root.setdefault('entries', {})
    tm_root = load_json(tm_path(a.book), {'schema': 1, 'book': meta['book'], 'entries': {}})
    tm = tm_root.setdefault('entries', {})
    book = meta['book']
    errors, warnings = [], []

    # مدخل‌های تازه‌ی دیکشنری
    new_lex = resp.get('newLexicon') or {}
    for lid, e in new_lex.items():
        missing = [f for f in LEX_REQUIRED if not e.get(f)]
        if missing:
            errors.append(f'مدخلِ «{lid}» این فیلدها را ندارد: {", ".join(missing)}')
        if lid in lex:
            warnings.append(f'مدخلِ «{lid}» از قبل در دیکشنری هست؛ نسخه‌ی موجود حفظ شد.')

    got = {t.get('id'): t for t in resp.get('translations', [])}
    for pid in meta['items']:
        if pid not in got:
            errors.append(f'{pid}: ترجمه ندارد.')
    for pid in got:
        if pid not in meta['items']:
            warnings.append(f'{pid}: در این دسته نیست؛ نادیده گرفته شد.')

    staged = {}
    for pid, info in meta['items'].items():
        t = got.get(pid)
        if not t:
            continue
        src = info['src']
        fa = (t.get('fa') or '').translate(NON_WESTERN_DIGITS)
        ar = (t.get('ar') or '').translate(NON_WESTERN_DIGITS)
        if not fa or not ar:
            errors.append(f'{pid}: ترجمه‌ی فارسی یا عربی خالی است.')
        if src.count('\n') != fa.count('\n') or src.count('\n') != ar.count('\n'):
            errors.append(f'{pid}: تعدادِ خط‌ها (\\n) با متنِ اصلی یکی نیست.')
        # شماره‌ی جای خالی نباید در ترجمه بیاید؛ فقط هشدار، چون گاهی عدد واقعاً جزءِ جمله است.
        for num in blank_numbers(src):
            if num not in re.findall(r'\d+', clean_nums(src)) and re.search(rf'(?<!\d){num}(?!\d)', fa):
                warnings.append(f'{pid}: عددِ {num} پیش از جای خالی شاید شماره‌ی سؤال باشد ولی در ترجمه آمده؛ بررسی کنید.')
        clean = src.replace('{blk}', '').replace('{/blk}', '')
        terms = []
        for tid in t.get('terms') or []:
            entry = lex.get(tid) or new_lex.get(tid)
            if entry is None:
                warnings.append(f'{pid}: مدخلِ «{tid}» نه در دیکشنری هست نه در newLexicon؛ حذف شد.')
                continue
            if not any(form_in(clean, f) for f in entry.get('forms') or [entry.get('headword', tid)]):
                warnings.append(f'{pid}: هیچ‌کدام از forms «{tid}» در متنِ پاراگراف نیست؛ حذف شد.')
                continue
            terms.append(tid)
        entry_out = {'src': src, 'fa': fa, 'ar': ar, 'terms': terms}
        # G-8: ترجمه‌ی پاسخِ هر جای خالی (همان تعداد و ترتیبِ {blk}ها)
        n_blanks = len(BLANK_ANSWER.findall(src))
        blanks = t.get('blanks')
        if n_blanks:
            if blanks is None:
                warnings.append(f'{pid}: {n_blanks} جای خالی دارد ولی blanks ندارد؛ پنجره‌ی پاسخ ترجمه‌ی جمله را فقط برای تک‌جای‌خالی نشان می‌دهد.')
            elif not isinstance(blanks, list) or len(blanks) != n_blanks:
                errors.append(f'{pid}: blanks باید دقیقاً {n_blanks} آیتم داشته باشد.')
            else:
                entry_out['blanks'] = [
                    {'fa': (b.get('fa') or '').translate(NON_WESTERN_DIGITS),
                     'ar': (b.get('ar') or '').translate(NON_WESTERN_DIGITS)}
                    for b in blanks if isinstance(b, dict)
                ]
        staged[info['key']] = entry_out

    for w in warnings:
        print('⚠️ ', w)
    if errors:
        for e in errors:
            print('❌', e)
        sys.exit('ادغام انجام نشد. پاسخ را اصلاح کنید (یا از هوش مصنوعی بخواهید اصلاح کند) و دوباره merge بزنید.')

    used = {tid for v in staged.values() for tid in v['terms']}
    added_lex = 0
    for lid, e in new_lex.items():
        if lid in lex or lid not in used:
            continue
        lex[lid] = {f: e[f] for f in LEX_REQUIRED}
        lex[lid]['books'] = [book]
        added_lex += 1
    for tid in used:
        books = lex[tid].setdefault('books', [])
        if book not in books:
            books.append(book)
    tm.update(staged)

    if a.dry_run:
        print(f'(آزمایشی) {len(staged)} ترجمه و {added_lex} مدخلِ تازه اضافه می‌شد.')
        return
    save_json(tm_path(a.book), tm_root)
    save_json(lex_path, lex_root)
    done_dir = os.path.join(a.book, 'enrichment', 'batches', 'done')
    os.makedirs(done_dir, exist_ok=True)
    for suffix in ('.ids.json', '.prompt.md'):
        p = os.path.join(a.book, 'enrichment', 'batches', batch + suffix)
        if os.path.exists(p):
            os.replace(p, os.path.join(done_dir, batch + suffix))
    print(f'✅ {len(staged)} ترجمه و {added_lex} مدخلِ تازه‌ی دیکشنری اضافه شد ({batch}).')


def main():
    ap = argparse.ArgumentParser(description='غنی‌سازیِ کتاب‌ها (ترجمه و کلمات تعاملی)')
    sub = ap.add_subparsers(dest='cmd', required=True)
    s = sub.add_parser('status')
    s.add_argument('book')
    p = sub.add_parser('prepare')
    p.add_argument('book')
    p.add_argument('--unit')
    p.add_argument('--audio', action='store_true')
    p.add_argument('--max-items', type=int, default=80)
    p.add_argument('--from-pending', action='store_true')
    m = sub.add_parser('merge')
    m.add_argument('book')
    m.add_argument('batch')
    m.add_argument('response')
    m.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()
    {'status': cmd_status, 'prepare': cmd_prepare, 'merge': cmd_merge}[a.cmd](a)


if __name__ == '__main__':
    main()
