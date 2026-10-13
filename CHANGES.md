# تغییرات — اسکریپت‌های صوتیِ هر Unit در enrich.py

فایل‌های تغییرکرده:
- `tools/enrichment/dump_docx.py`: نگاشتِ «فایلِ صوتی ← Unit» از لینک‌های صوتیِ docxِ کتاب
  (`AUDIO_UNITS`)
- `tools/enrichment/enrich.py`:
  - `prepare --audio --unit` فقط اسکریپت‌های همان Unit را می‌آورد.
  - `status` Unitِ هر فایلِ صوتی را نشان می‌دهد.
- `tools/enrichment/ENRICHMENT-GUIDE.md`: بخشِ ۴-۶ (اسکریپت‌های صوتی)
- `WORKFLOW.md`: روالِ «هر Unit همراهِ اسکریپت‌های صوتی‌اش» (هم‌سان با نسخه‌ی Flutter)

تغییرِ هم‌زمان در ielts_assistant: ترجمه‌ی کاملِ Mindset 1، Unit 1، شاملِ متن و ۱۳ اسکریپتِ
صوتی، و ۷۳ مدخلِ تازه‌ی دیکشنری.

## مشکل
فایلِ اسکریپت‌های صوتی هیچ نشانی از Unit ندارد: فقط جدول‌هایی است با نامِ فایل‌های
`Mindset_L1_02.mp3` تا `Mindset_L1_71.mp3`. پس `prepare --audio --unit "UNIT/01"` چیزی پیدا
نمی‌کرد، چون فیلتر روی نامِ فایلِ صوتی اعمال می‌شد.

## راه‌حل
هر فایلِ صوتی در docxِ کتاب یک هایپرلینکِ `audio:Mindset_L1_02.mp3` دارد؛ همان چیزی که
`AudioLinksIndex`ِ اکسترکتور هم از آن ساخته می‌شود. حالا:
- `dump_docx.py` هنگامِ خواندنِ کتاب، برای هر فایلِ صوتی ثبت می‌کند که اولین لینکش در کدام Unit
  است.
- `enrich.py` با همین نگاشت، اسکریپت‌های یک Unit را جدا می‌کند.

نتیجه برای Mindset 1:

| Unit | فایل‌های صوتی |
|---|---|
| 1 | 02 تا 14 |
| 2 | 15 تا 25 |
| 3 | 26 تا 31 |
| 4 | 32 تا 39 |
| 5 | 40 تا 48 |
| 6 | 49 تا 58 |
| 7 | 59 تا 65 |
| 8 | 66 تا 71 |

روالِ هر Unit:

```bat
python enrich.py prepare "…\Mindset 01" --unit "UNIT/02"
python enrich.py prepare "…\Mindset 01" --unit "UNIT/02" --audio
```

## تست
روی Mindset 1 اجرا شد:
- `status` کنارِ هر فایلِ صوتی Unitِ آن را نشان داد.
- `prepare --unit "UNIT/01" --audio` یک دسته با ۱۳ فایلِ صوتیِ Unit 1 (۷۱ متنِ یکتا) ساخت. همین
  دسته ترجمه و ادغام شد.
