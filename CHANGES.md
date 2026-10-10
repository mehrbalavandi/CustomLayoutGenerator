# تغییرات — هندسه‌ی شماره‌ی لیست برای تورفتگیِ دقیقِ Word

فایل‌های تغییرکرده: `MainForm.cs`، `Numberingresolver.cs`، `Models.cs`
تغییرِ هم‌زمان در فلاتر: مخزنِ ielts_assistant (قاعده‌ی تورفتگیِ Word؛ جزئیات و
آمار در CHANGES.mdِ همان مخزن)

## چرا
تورفتگیِ لیست‌ها در اپ بیشتر از Word بود. فلاتر حالا شروعِ متنِ خطِ اول را با
قاعده‌ی خودِ Word حساب می‌کند: اگر شماره در hanging جا شود، متن از IndentLeft؛
وگرنه از tab stopِ بعدی. برای این تصمیم، فلاتر باید عرضِ شماره را دقیقاً مثلِ Word
بداند و بداند بعد از شماره چه می‌آید.

## فیلدهای جدیدِ پاراگراف (همه اختیاری؛ فقط وقتی با پیش‌فرض فرق دارند)

| فیلد | از کجا | پیش‌فرض |
|---|---|---|
| `ListMarkerSize` | اندازه‌ی فونتِ شماره (pt): سطحِ numbering ← نشانه‌ی پایانِ پاراگراف ← استایلِ کاراکتریِ نشانه ← استایلِ پاراگراف (یا Normal) ← docDefaults | اندازه‌ی اولین اسپن |
| `ListMarkerScale` | فشردگیِ افقی (`w:w`): سطحِ numbering ← نشانه‌ی پایانِ پاراگراف؛ مثلاً ۰.۸۴ | ۱ |
| `ListMarkerAlign` | `w:lvlJc` = right/center (start/end هم نگاشت می‌شوند) | left |
| `ListSuffix` | `w:suff` = space/nothing | tab |
| `ListTabStop` | `w:defaultTabStop` سند (pt)؛ فقط اگر ≠ ۳۶ | ۳۶ |

- فشردگیِ افقی لازم بود: در Mindset 2 شماره‌ی «10:» با ۸۴٪ در hangingِ ۱۴.۲pt جا
  می‌شود، ولی بدونِ آن نه.
- `NumberingResolver` دو متدِ جدید دارد: `LevelSuffix` و `LevelJustification`.
  rPrِ سطحِ numbering هم با `LevelRunProperties` در دسترس است.
- `CloneParagraphProperties` همه‌ی این فیلدها را کپی می‌کند.

## بررسی روی سه کتاب
- IndentLeft و IndentFirstLineِ استخراج‌شده در **همه‌ی** پاراگراف‌های شماره‌دار با
  مقدارِ مؤثرِ Word (استایل ← سطحِ numbering ← مستقیم) برابر بود؛ تغییری لازم نشد.
- همه‌ی لیست‌ها `suff=tab` و `defaultTabStop=36pt` دارند، پس `ListSuffix` و
  `ListTabStop` در این کتاب‌ها نوشته نمی‌شوند. برای کتاب‌های آینده آماده‌اند.
- ۱۲ سطح در Mindset 1 راست‌چین‌اند (ص۸۲).

## نیاز به استخراجِ مجدد
بله، هر سه کتاب.

## تست
SDKِ .NET در محیطِ من در دسترس نبود. هر سه فایل با پارسرِ C# (tree-sitter) بدونِ
خطای نحوی پارس شدند. نام و نوعِ کلاس‌های OpenXml (`CharacterScale.Val` = IntegerValue،
`DefaultTabStop.Val` = Int16Value، `Level.LevelSuffix` و `Level.LevelJustification`)
با سورسِ خودِ SDK چک شد. لطفاً یک Build بزنید. هر سه فایل `MainForm.cs`،
`Numberingresolver.cs` و `Models.cs` با هم عوض شوند.
