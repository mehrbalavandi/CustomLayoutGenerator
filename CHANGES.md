# تغییرات — فهرستِ ساختاریِ کتاب (Outline) از روی سطحِ ساختارِ Word

فایل‌های تغییرکرده:
- `Models.cs`
- `MainForm.cs`
- `Bookoutputwriter.cs`
- `RULES-extraction.md` (قانونِ O-4 + «راهنمای نویسنده‌ی سند»)
- `ARCHITECTURE-extractor.md`

تغییرِ هم‌زمان در مخزنِ ielts_assistant:
- ناوبرِ Unit/مهارت در اپ.
- علامت‌گذاریِ سه فایلِ docxِ Mindset.

## هدف
- کاربر حینِ مطالعه بداند در کدام Unit و کدام مهارت است.
- با یک لمس به Unit یا مهارتِ دیگر برود.

## چرا «سطحِ ساختار» و نه استایلِ مخصوصِ Unit/مهارت
هر کتاب ساختارِ خودش را دارد (Unit/مهارت، فصل/بخش، Test/Part، …). پس اکسترکتور
اسمِ استایل یا متنِ خاصی را نمی‌شناسد. فقط **سطحِ ساختارِ خودِ Word** (Outline
Level) را می‌خواند، یعنی همان چیزی که Navigation Pane و فهرستِ مطالبِ Word از آن
ساخته می‌شوند:
- هر استایلی که Outline Level داشته باشد کار می‌کند: «Nav 1»/«Nav 2»، «Heading 1/2»،
  یا سطحی که مستقیم در Paragraph ← Outline level داده شده.
- **چک کردن در خودِ Word:** View ← Navigation Pane همان درختی را نشان می‌دهد که اپ
  نشان خواهد داد.
- کتابِ تازه هیچ تغییری در کد نمی‌خواهد.

## چه چیزی اضافه شد
1. **`ParagraphData.OutlineLevel`:**
   - سطحِ مؤثرِ پاراگراف: مستقیم، وگرنه از زنجیره‌ی استایل.
   - فقط روی اولین تکه‌ی هر پاراگرافِ Word گذاشته می‌شود.
   - در JSONِ صفحه نوشته **نمی‌شود** (`[JsonIgnore]`)، پس هَشِ صفحه‌ها عوض نمی‌شود و
     کاربرها صفحه‌ها را دوباره دانلود نمی‌کنند.
   - در `CloneParagraphProperties` هم کپی می‌شود.
2. **`Outline` در index.json** (درخت، به ترتیبِ سند):
   ```json
   "Outline": [
     { "Title": "UNIT/01: THE MAN-MADE ENVIRONMENT", "Level": 1,
       "PageNumber": 1, "ParaIndex": 0, "AtPageTop": true,
       "Children": [
         { "Title": "READING", "Level": 2, "PageNumber": 1, "ParaIndex": 1, "AtPageTop": true },
         { "Title": "WRITING", "Level": 2, "PageNumber": 7, "ParaIndex": 1, "AtPageTop": true },
         …
       ] },
     …
   ]
   ```
   - `ParaIndex`: همان قراردادِ `AudioLinksIndex`، یعنی اندیسِ پاراگرافِ بیرونیِ صفحه.
     تیترِ داخلِ سلولِ جدول هم حساب می‌شود.
   - `AtPageTop`: اگر قبل از تیتر در همان صفحه روی‌هم حداکثر ۸۰ نویسه‌ی دیدنی آمده
     باشد و هیچ عکسی نباشد (مثلِ بنرِ «UNIT/01: …»)، پرش یعنی رفتن به بالای صفحه.
     وگرنه اپ خودِ پاراگرافِ تیتر را نشان می‌دهد.
   - **عنوان:** متنِ دیدنیِ پاراگراف. رانِ کوچک‌تر از ۲pt کنار گذاشته می‌شود. در Mindset 1
     سه تیترِ «READING» یک کپیِ نامرئیِ ۱pt از خودشان دارند.
   - **تیترِ تکراری:** تیترِ هم‌سطح و هم‌عنوان با گره‌ی بازِ فعلی (بنرِ تکراری) گره‌ی
     تازه نمی‌سازد.
   - **سازگاری:** کتابی که هیچ تیترِ سطح‌دار ندارد، `Outline` ندارد. نسخه‌های قدیمیِ اپ
     این فیلد را نادیده می‌گیرند. `SchemaVersion` همان ۲ می‌ماند، چون تغییر فقط اضافه
     است.

## سه کتابِ Mindset (در مخزنِ ielts_assistant علامت‌گذاری شد)
- در هر کتاب: ۸ تیترِ Nav 1 (اولین بنرِ «UNIT/0X: …» هر Unit) و ۳۲ تیترِ Nav 2
  (READING/WRITING/LISTENING/SPEAKING).
- ترتیبِ مهارت‌ها در هر سه کتاب بررسی شد و همه R/W/L/S بودند.
- درختی که شبیه‌سازی از روی فایل‌ها ساخت (شماره‌ی صفحه با شروعِ ۱):
  - Mindset 1: Unit 1 ← R@1، W@5، L@9، S@12؛ … Unit 8@139.
  - Mindset 2: Unit 1 ← R@1، W@7، L@14، S@17؛ … Unit 8@141.
  - Mindset 3: Unit 1 ← R@1، W@7، L@12، S@17؛ … Unit 8@152.

## نیاز به استخراجِ مجدد
بله، هر سه کتاب، تا `Outline` در index.json نوشته شود. هَشِ صفحه‌ها عوض نمی‌شود.

## تست
SDKِ .NET در محیطِ من در دسترس نبود.
- سه فایل با پارسرِ C# (tree-sitter) بدونِ خطای نحوی پارس شدند.
- API‌های OpenXML (`ParagraphProperties.OutlineLevel`،
  `StyleParagraphProperties.OutlineLevel`، `Val: Int32Value`) با سورسِ SDK چک شدند.

لطفاً یک Build بزنید.
