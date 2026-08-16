# معماری CustomLayoutGenerator (استخراج‌کننده‌ی سی‌شارپ)

> این سند برای وقتی است که مدتی از پروژه دور بوده‌اید و می‌خواهید سریع بفهمید
> «چی به چیه». وضعیت مخزن: کامیت `b43151b`.

## این برنامه چه‌کار می‌کند

یک اپِ WinForms تک‌دکمه‌ای است که یک فایل ورد را می‌گیرد و آن را به مجموعه‌ای
از فایل‌های JSON تبدیل می‌کند که اپِ فلاتر می‌خواند. **کلِ منطقِ فهمیدنِ ورد
این‌جاست؛ فلاتر فقط JSON را رندر می‌کند.** هر وقت چیزی در اپ «اشتباه نشان
داده می‌شود»، اول باید بپرسید: آیا اصلاً در JSON درست آمده؟

جای این برنامه در زنجیره:

```
سندِ ورد ──► CustomLayoutGenerator (این‌جا) ──► پوشه‌ی JSON
                                                    │
                                                    ▼
                                        بک‌اندِ لاراول (bookassistant-backend)
                                                    │  ZIP + نسخه‌بندی
                                                    ▼
                                            اپِ فلاتر (ielts_assistant)
```

دو جریانِ ورودی جدا وجود دارد: **سندِ کتاب** (صفحات) و **سندِ اسکریپتِ صوت**
(یک جدولِ دوردیفه به‌ازای هر فایل صوتی، با زمان‌بندیِ کلمه‌به‌کلمه).

## نقشه‌ی فایل‌ها

| فایل | خطوط | مسئولیت |
|---|---|---|
| `Program.cs` | 16 | نقطه‌ی ورودِ WinForms. کاری نمی‌کند. |
| `MainForm.Designer.cs` | 59 | فرمِ طراحی‌شده (یک دکمه). |
| **`MainForm.cs`** | **2724** | **کلِ مغزِ برنامه.** پایین‌تر تفکیک شده. |
| `Models.cs` | 191 | شکلِ دیتا. قراردادِ بین سی‌شارپ و فلاتر. |
| `Responsivelowering.cs` | 280 | تبدیلِ نامِ استایلِ جدولِ ورد به رفتارِ ریسپانسیو. |
| `Numberingresolver.cs` | 208 | شماره‌گذاریِ خودکارِ لیست‌ها (`numbering.xml`). |
| `FontResolver.cs` | 449 | تشخیصِ فونتِ مؤثرِ هر ران. |
| `Bookoutputwriter.cs` | 214 | نوشتنِ فایل‌ها روی دیسک + نسخه‌بندی. |
| `Class1.cs` | 4 | خالی. بازمانده‌ی قالبِ پروژه؛ می‌شود حذفش کرد. |

---

## Models.cs — قراردادِ داده

اگر بخواهید بفهمید اپِ فلاتر چه می‌بیند، **از این‌جا شروع کنید**. هر
پراپرتی این‌جا مستقیماً یک کلید در JSON می‌شود (با `NullValueHandling.Ignore`،
پس فیلدهای null اصلاً نوشته نمی‌شوند — به همین دلیل افزودنِ فیلدِ جدید
همیشه با کتاب‌های قدیمی سازگار است).

سلسله‌مراتب:

```
PageData          ← یک صفحه (PageNumber + Paragraphs)
 └ ParagraphData  ← یک پاراگراف
    ├ ویژگی‌های بلوکی: Alignment, Direction, IndentLeft/Right/FirstLine,
    │  SpaceBefore/After, LineSpacing, FillColor
    ├ لیست: ListType, ListLevel, ListMarker, ListMarkerBold, ListMarkerColor,
    │  KeepListMarkerVisible
    ├ صوت: StartMs, EndMs, AudioTrackName
    └ Spans
       └ SpanData ← یک تکه‌ی درون‌خطی. Type یکی از:
          "text"   → Content + Markers + Url + رنگ/زیرخط/فاصله‌ی حروف
          "image"  → مسیر + ImageWidth/Height + FloatPosition
          "table"  → TableRows (که خودشان Cell دارند و هر Cell پاراگراف)
          "layout" → همان جدول، ولی Responsivelowering نوعش را عوض کرده
```

`Markers` یک لیستِ رشته‌ای است و قالب‌بندیِ ران را نگه می‌دارد:
`b`, `i`, `u`, `s`, `sub`, `sup`, `smallcaps`, و همچنین `fn:<نام فونت>` و
`pindent:<pt>`.

**نکته‌ی مهم:** `SpanData` هم برای متن، هم عکس، هم جدول استفاده می‌شود —
یعنی یک کلاسِ بزرگ با فیلدهایی که فقط برای بعضی نوع‌ها معنا دارند. عمدی است
و ساده‌ترین راه برای JSONِ یکدست بود.

## MainForm.cs — تفکیکِ مسئولیت‌ها

بزرگ است، ولی چند ناحیه‌ی مشخص دارد:

### ۱) جریانِ اصلی
- `btnSelectFile_Click` (خط ۱۱۹) — کلِ سناریو: انتخابِ فایل ← پرسیدنِ شماره‌ی
  صفحه‌ی شروع ← کپیِ موقت ← پردازش ← ادغام‌های BlankWord ← نوشتنِ خروجی.
- `PromptForStartPage` (۷۲) — دیالوگِ دستیِ راست‌به‌چپ برای شماره‌ی شروع.
- `ProcessWordDocument` (۲۱۳) — پیمایشِ بدنه‌ی سند، شکستنِ صفحات روی
  page-break، ساختِ `PageData`.

### ۲) تبدیلِ پاراگراف و ران
- `ParseParagraph` — ویژگی‌های بلوکی + شماره‌ی لیست + جهت + هم‌ترازی.
- `ProcessRun` (۱۰۷۳) — قلبِ استخراجِ متن: ساختِ `SpanData` برای هر ران،
  شاملِ رنگ، پس‌زمینه، زیرخط، فاصله‌ی حروف، هایپرلینک، عکس.
- `ExtractRunMarkers` (۲۲۲۸) — تولیدِ لیستِ `Markers`.

### ۳) حل‌کردنِ قالب‌بندی (مهم‌ترین دسته‌ی باگ‌های تاریخی)
ورد قالب‌بندی را در سه لایه نگه می‌دارد و همه باید با هم دیده شوند:

```
rPrِ مستقیمِ ران  →  استایلِ کاراکتری (+ زنجیره‌ی BasedOn)  →  استایلِ پاراگراف
```

- `StyleRunPropsChain` (۲۱۵۹) — پیمایشِ زنجیره‌ی `BasedOn`.
- `ResolveRunProp<T>` (۲۱۷۸) — نسخه‌ی عمومی؛ با `GetFirstChild<T>()` روی هر
  دو کلاس کار می‌کند. **هر خاصیتِ تازه‌ای از `rPr` که بعداً لازم شد، فقط یک
  خط با نامِ کلاسِ عنصر می‌خواهد.**
- `IsBold` / `IsItalic` / `IsAllCaps` / `GetFontSizeFromStyleId` /
  `GetColorFromStyleId` / `GetAlignmentFromStyleId` — همان کار برای موارد
  قدیمی‌تر.
- `PrefersComplexScript` / `ContainsComplexScriptChar` (۱۹۵۸، ۱۹۷۴) —
  گاردهای Complex Script. ورد در سندِ دوزبانه `bCs`/`iCs` را تقریباً روی
  همه‌چیز می‌پاشد؛ بدونِ این گاردها متنِ انگلیسی به‌غلط ایتالیک می‌شود.
- `MapUnderline` (۱۷۷۸) / `MapHighlight` (۲۲۰۳) / `MapAlignment` (۱۶۳۱) —
  نگاشتِ واژگانِ ورد به واژگانِ فلاتر. **این نگاشت‌ها عمداً این‌طرف‌اند** تا
  فقط یک جا نگهداری شوند.

### ۴) جدول‌ها
- `ParseTable` (۱۲۹۲) — ساختِ `SpanData` از نوع table، بازگشتی برای
  جدولِ تودرتو.
- `ExtractCellProperties` (۲۶۴۴) + `ExtractSmartCellPadding` (۲۶۷۸) —
  بوردر، رنگ، عرض، هم‌ترازیِ عمودی، پدینگ.
- `TableStyleConditionalBold` (۱۲۶۵) — بولدِ ردیفِ اول که از
  `<w:tblStylePr>` می‌آید نه از خودِ ران.

### ۵) جای‌خالی‌ها (BlankWord)
سه استایلِ ورد که متن را مخفی می‌کنند و در اپ آیکونِ چشم می‌شوند:

| استایل | رفتار |
|---|---|
| BlankWord1 | یک ران درون‌خطی. `MergeConsecutiveInlineBlanks` (۱۷۱۰) ران‌های پشتِ‌هم را یکی می‌کند تا یک چشم به‌جای چند تا. |
| BlankWord2 | کلِ پاراگراف؛ پاراگراف‌های پشتِ‌هم در یک بلوک ادغام می‌شوند. `MergeBlankWord2Paragraphs` (۴۳۷). |
| BlankWord3 | کلِ پاراگراف ولی **بدونِ** ادغام و شماره‌ی لیست بیرون می‌ماند. `WrapBlankWord3Paragraphs` (۶۰۹). |

متنِ مخفی با `{blk}…{/blk}` علامت می‌خورد. `CloneSpan` (۱۸۱۸) و
`CloneParagraphProperties` (۱۶۶۰) موقعِ این ادغام‌ها استفاده می‌شوند —
**اگر فیلدِ جدیدی به `SpanData` اضافه کردید، حتماً به `CloneSpan` هم اضافه‌اش
کنید**، وگرنه در جای‌خالی‌ها بی‌صدا گم می‌شود.

## Responsivelowering.cs

بعد از استخراج اجرا می‌شود و **نامِ استایلِ جدولِ ورد** را به دو فیلدِ
اعلانی ترجمه می‌کند که فلاتر می‌فهمد:

- `BorderMode`: `all` | `outer` | `inner` | `cell` | `none` | `firstRowOuter` | `outerThickFirstRow`
- `WidthMode`: `content` | `equal` | `proportional` | `fill` | `natural`

به‌علاوه `ResponsiveStrategy`, `LayoutDirection`, `LayoutReflow`.

استایل‌های شناخته‌شده: DottedTable، BorderedTable، CompactTable، OutsideTable،
HeaderOutsideTable، FigureTable، HBTable، NormalTable، TipTable،
ColumnStackTable، MultiColumnTable، CommonTable. **`default:`** (هر استایلِ
ناشناخته) به `cell`/`natural` می‌رود، یعنی «همان‌طور که در ورد است».

**برای افزودنِ رفتارِ جدید، معمولاً فقط یک `case` این‌جا لازم است** — نه
تغییر در فلاتر.

## Numberingresolver.cs

شماره‌گذاریِ خودکارِ لیست‌ها را از `numbering.xml` می‌خواند و شمارنده‌ها را
نگه می‌دارد. نکته‌ی غیرشهودی: **قالب‌بندیِ شماره (بولد/رنگ) در تعریفِ سطحِ
numbering است، نه در رانِ متن** — `LevelBold` و `LevelColor` برای همین‌اند.

## FontResolver.cs

فونتِ مؤثرِ هر ران را از `rFonts` (Ascii در برابر ComplexScript)، استایل، و
`docDefaults` حل می‌کند. `ContainsPersianOrArabic` تصمیم می‌گیرد کدام نسخه
برنده شود.

## Bookoutputwriter.cs

خروجی را می‌نویسد:

```
index.json          ← فهرستِ صفحات + AudioScripts + Interactives
pages/page_0001.json
audio_scripts/…
images/  audio/
```

`ShortHash` هَشِ محتوا را می‌سازد که در `Version` هر صفحه می‌نشیند —
**همین است که دانلودِ دلتایی را ممکن می‌کند**؛ صفحه‌ای که عوض نشده دوباره
دانلود نمی‌شود. نامِ فایلِ صفحه از `PageNumber` می‌آید، پس شماره‌ی شروعِ
دلخواه خودبه‌خود در نامِ فایل‌ها هم منعکس می‌شود.

## تله‌های شناخته‌شده (برای خودِ آینده)

1. **`.Val?.Value.ToString()` روی enumهای OpenXML رشته‌ی بی‌معنی می‌دهد**
   (مثل `"bordervalues { }"`). همیشه `.Val?.InnerText`. این باگ تا حالا در
   بوردر، هم‌ترازیِ عمودی، هم‌ترازیِ جدول و vMerge پیدا شده. اگر خاصیتی
   «کار نمی‌کند»، اول دنبالِ این الگو بگردید.
2. **`StyleRunProperties` برای همه‌ی فرزندانش پراپرتیِ نوع‌دار ندارد** (مثلاً
   `.Highlight` ندارد) — به‌جایش `GetFirstChild<T>()`.
3. **قالب‌بندی ممکن است از استایل بیاید نه از ران.** هر بار که چیزی «فقط
   بعضی جاها کار می‌کند»، احتمالش زیاد است.
4. **در `rPr`، کلاسِ `w:spacing` همان `Spacing` است** (فاصله‌ی حروف)؛
   `SpacingBetweenLines` مالِ `pPr` است.
5. فیلدِ جدید در `SpanData` ⇒ اضافه‌کردن به `CloneSpan` را فراموش نکنید.

## چیزهایی که عمداً پیاده نشده‌اند

`w:vanish` (متنِ مخفیِ خودِ ورد — چون ممکن است استایل‌های BlankWord از آن
استفاده کنند و متن پاک شود)، `w:position`، `w:w`، `w:em`، outline/shadow،
vMerge، colspanِ بومی. `w:dstrike` به خط‌خوردگیِ ساده تبدیل می‌شود چون
فلاتر خط‌خوردگیِ دوتایی ندارد.
