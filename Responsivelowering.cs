using System.Collections.Generic;

namespace CustomLayoutGenerator
{
    /// <summary>
    /// نام‌های استایلِ جدولِ Word را به primitiveهای declarativeِ ریسپانسیو «تبدیل»
    /// می‌کند تا فلاتر به‌جای switch روی نام استایل، فقط با چند فیلدِ عمومی کار کند.
    /// افزودن یک استایلِ جدید = یک case در این switch (بدون تغییر در فلاتر).
    /// روی spanها به‌صورت بازگشتی کار می‌کند (جدولِ تودرتو، innerSpans و سلول‌ها).
    ///
    /// 🐞 بازطراحیِ کامل: قبلاً فقط ResponsiveStrategy (برایِ عرض/اسکرول) و
    /// Borders.Val (برایِ نمایش/عدم‌نمایشِ بوردر) ست می‌شد، و فلاتر برایِ
    /// تصمیم‌های ظریف‌تر (مثلاً «فقط بوردرِ بیرونی» برایِ OutsideTable) مجبور
    /// بود مستقیماً نامِ استایل را چک کند (isOutsideTable, isBorderedTable, ...)
    /// — یعنی هر رفتارِ جدید نیاز به یک پرچمِ جدید در فلاتر داشت. حالا دو فیلدِ
    /// صریح و توسعه‌پذیر اضافه شده: BorderMode (کدام لبه‌ها بوردر دارند) و
    /// WidthMode (عرض چطور تعیین می‌شود). فلاتر فقط رویِ همین دو مقدار سوییچ
    /// می‌کند؛ هر حالتِ آینده (مثلاً «فقط بوردرِ داخلی» یا «فقط ردیفِ اول») یعنی
    /// فقط یک مقدارِ جدید این‌جا + یک caseِ متناظر در فلاتر.
    /// </summary>
    public static class ResponsiveLowering
    {
        public static void Apply(List<PageData> pages)
        {
            foreach (var page in pages)
                foreach (var para in page.Paragraphs)
                    LowerParagraph(para);
        }

        private static void LowerParagraph(ParagraphData para)
        {
            if (para?.Spans == null) return;
            foreach (var span in para.Spans) LowerSpan(span);
        }

        private static void LowerSpan(SpanData span)
        {
            if (span == null) return;

            // بازگشت: اسپن‌های داخلی و پاراگراف‌های داخل سلول‌های جدول
            foreach (var inner in span.InnerSpans) LowerSpan(inner);
            foreach (var row in span.TableRows)
                foreach (var cell in row.Cells)
                    foreach (var p in cell.Paragraphs)
                        LowerParagraph(p);

            if (span.Type != "table") return;

            var styleKey = span.TableStyleName ?? span.TableStyleId ?? "";
            switch (styleKey)
            {
                case "MultiColumnTable":
                    // 🌟 متنِ چندستونی: در صفحهٔ عریض همهٔ ستون‌ها کنارِ هم و
                    // داخلِ یک جدول‌اند که فقط بوردرِ بیرونی‌اش دیده می‌شود
                    // (BorderMode="outer" — خطوطِ داخلیِ بینِ ستون‌ها رسم
                    // نمی‌شود)؛ در صفحهٔ باریک، همان استراتژیِ "stack" فعال
                    // می‌شود و هر ستون به یک جدولِ تک‌ستونهٔ بوردردارِ مستقل
                    // تبدیل می‌گردد. یعنی ترکیبی از رفتارِ OutsideTable
                    // (عریض) و ColumnStackTable (باریک) — بدونِ نیاز به
                    // مکانیزمِ تازه، فقط با کنارِ هم گذاشتنِ همان دو پرچمِ
                    // موجود.
                    span.ResponsiveStrategy = "stack";
                    span.Type = "layout";
                    span.LayoutDirection = "row";
                    span.LayoutReflow = "stack";
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
                    span.BorderMode = "outer";
                    span.WidthMode = "equal";
                    break;

                case "ColumnStackTable":
                    // جدولِ چیدمانی → نودِ layoutِ ستون‌محور که در صفحهٔ کوچک عمودی می‌شود
                    span.ResponsiveStrategy = "stack";
                    span.Type = "layout";
                    span.LayoutDirection = "row";
                    span.LayoutReflow = "stack";
                    break;

                case "DottedTable":
                    span.ResponsiveStrategy = "collapseToCards";
                    span.Borders = span.Borders ?? new BorderDetail();
                    span.Borders.Val = "dotted";
                    // 🐞 عمداً "none": ظاهرِ نقطه‌چین/کارتی فقط برایِ حالتِ
                    // موبایلِ collapseToCards استفاده می‌شود، نه یک خطِ
                    // واقعیِ رسم‌شده در نمایِ گریدِ عادی — قبلاً اشتباهاً
                    // "all" بود که باعث می‌شد همه‌ی DottedTableها بوردر
                    // نشان بدهند، درحالی‌که قبلاً هیچ‌وقت نشان نمی‌دادند.
                    span.BorderMode = "none";
                    span.WidthMode = "equal";
                    break;

                case "BorderedTable":
                    span.ResponsiveStrategy = "horizontalScroll";
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
                    span.BorderMode = "all";
                    span.WidthMode = "fill";
                    break;

                case "CompactTable":
                    // 🐞 برای جدول‌های تک‌سلولیِ کوچک مثلِ «شمارهٔ تمرین» (۰۱،
                    // ۰۲، ...) که باید یک جعبهٔ کوچکِ بوردردار و
                    // متناسب‌با‌محتوا باشند، نه یک جدولِ عریض/اسکرول‌شونده.
                    // WidthMode="content" یعنی فلاتر هیچ‌وقت این را کش
                    // نمی‌آورد؛ دقیقاً به‌اندازهٔ متنِ داخلش می‌ماند.
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
                    span.BorderMode = "all";
                    span.WidthMode = "content";
                    break;

                case "FigureTable":
                    // 🐞 برای جدولِ تک‌سلولیِ «عکس + زیرنویسِ رنگی» (مثلِ نمودار
                    // با چند منحنیِ رنگی که زیرش توضیحِ رنگ‌ها می‌آید): چون
                    // هر دو پاراگراف (عکس و زیرنویس) در یک سلول‌اند، اگر
                    // فلاترْ عرضِ لازمِ این جدول را بر اساسِ عرضِ طبیعیِ عکسِ
                    // داخلش حساب کند (نه فقط تعدادِ ستون)، عکس و زیرنویس با
                    // هم و در یک عرضِ یکسان اسکرول/رندر می‌شوند. برخلافِ
                    // BorderedTable، اینجا بوردرِ پیش‌فرض اضافه نمی‌شود — یک
                    // جعبه‌ی عکس/زیرنویس معمولاً بوردر نمی‌خواهد.
                    span.ResponsiveStrategy = "horizontalScroll";
                    span.BorderMode = "none";
                    span.WidthMode = "fill";
                    break;

                case "HBTable":
                    // 🐞 "Horizontal and Bordered Table": مثلِ BorderedTable
                    // (اسکرولِ افقی + بوردرِ پیش‌فرض)، ولی به‌عنوانِ استایلِ
                    // اسمِ‌جداگانه‌ای که کاربر مستقیماً برای همین منظور در
                    // Word ساخته — تا با نامِ روشن‌تری از BorderedTableِ قدیمی
                    // جدا باشد.
                    span.ResponsiveStrategy = "horizontalScroll";
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
                    span.BorderMode = "all";
                    span.WidthMode = "fill";
                    break;

                case "CommonTable":
                    // 🐞 استایلِ عمومیِ وفادار: جدول دقیقاً همان‌طور که در سند
                    // است رندر می‌شود — بوردر و رنگِ پس‌زمینه در سطحِ *سلول*
                    // (نه کلی) و عیناً. هیچ بوردری اجباری اضافه نمی‌شود؛ هر سلول
                    // بوردر/رنگِ خودش را از سند دارد (Borders سلول از قبل توسطِ
                    // ExtractSmartCellBorders و FillColor استخراج شده‌اند). عرضِ
                    // ستون‌ها از عرضِ مطلقِ سند (WidthPt) می‌آید و تصمیمِ اسکرولِ
                    // افقی در فلاتر با مقایسه‌ی مجموعِ عرضِ طبیعی با عرضِ دستگاه
                    // گرفته می‌شود (نه یک پرچمِ ثابت). همه‌ی سلول‌ها — حتی خالی —
                    // رسم می‌شوند. ColumnStack/MultiColumn/Dotted و بقیه‌ی
                    // استایل‌های قبلی دست‌نخورده و با همان کاربردِ قبلی می‌مانند.
                    span.BorderMode = "cell";
                    span.WidthMode = "natural";
                    break;

                case "NormalTable":
                    // 🐞 درخواستِ کاربر: بوردرِ جدول فقط وقتی رسم شود که سندِ
                    // اصلی واقعاً بوردر داشته باشد و آن بوردر solid باشد (نه
                    // none/dotted/dashed). قبلاً NormalTable بی‌قیدوشرط
                    // BorderMode="all" می‌گرفت، برای همین جدولِ کاملاً بی‌بوردرِ
                    // تمرینِ ۰۸ صفحه‌ی ۲۶ (Borders خالی هم در سطحِ جدول هم در
                    // سطحِ سلول‌ها) هم بوردر می‌گرفت — دقیقاً همان نقضِ قاعده‌ای
                    // که کاربر گزارش کرد. حالا وقتی هیچ بوردرِ solidی در سند
                    // نباشد BorderMode="none" می‌شود. بدونِ اسکرولِ افقی می‌ماند.
                    {
                        bool normalHasSolid = HasSolidBorderInSource(span);
                        span.Borders = span.Borders ?? new BorderDetail();
                        span.BorderMode = normalHasSolid ? "all" : "none";
                        span.WidthMode = "equal";
                    }
                    break;

                case "TipTable":
                    // 🐞 نه اسکرولِ افقی، نه بوردرِ اجباری — رفتارِ خنثی/عادی؛
                    // اگر خودِ سند بوردر دارد همان حفظ می‌شود، وگرنه چیزی
                    // اضافه نمی‌شود. BorderMode/WidthMode عمداً ست نمی‌شوند
                    // (null می‌مانند) تا فلاتر به دادهٔ خامِ per-cell رجوع کند.
                    break;

                case "OutsideTable":
                    // 🐞 اسکرولِ افقی لازم دارد (اگر واقعاً محتوا جا نشود)،
                    // ولی فقط بوردرِ دورتادورِ کلِ جدول باید دیده شود، نه
                    // خطوطِ داخلیِ بینِ سلول‌ها/ردیف‌ها. WidthMode="proportional"
                    // یعنی هر ستون متناسب با محتوایِ خودش عرض می‌گیرد (نه
                    // مساوی با بقیه)، تا وقتی محتوا کوتاه است له نشود و وقتی
                    // واقعاً جا نیست به‌درستی اسکرول بگیرد.
                    span.ResponsiveStrategy = "horizontalScroll";
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
                    span.BorderMode = "outer";
                    span.WidthMode = "proportional";
                    break;

                case "HeaderOutsideTable":
                    // 🐞 درخواستِ کاربر: مثلِ OutsideTable (اسکرولِ افقی + عرضِ
                    // متناسب + فقط بوردرِ بیرونی)، ولی با دو تفاوت — (۱) بوردرِ
                    // بالای جدول هم رسم شود (BorderMode="outerThickFirstRow"
                    // مثلِ "outer" یک Border.all کاملِ چهارطرفه دورِ کل می‌کشد،
                    // پس بالای جدول هم می‌آید)، و (۲) خطِ زیرِ ردیفِ اول (جداکننده‌ی
                    // سرستون) ضخیم‌تر از بقیه‌ی خطوط دیده شود. رفتارِ عرض/اسکرول
                    // دقیقاً همان OutsideTable است.
                    span.ResponsiveStrategy = "horizontalScroll";
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
                    span.BorderMode = "outerThickFirstRow";
                    span.WidthMode = "proportional";
                    break;

                // 🐞 دو حالتِ جدید که کاربر برایِ آینده خواسته بود — از قبل
                // آماده‌اند تا وقتی در Word یک جدول را با یکی از این نام‌ها
                // استایل داد، بلافاصله کار کند، بدونِ نیاز به تغییرِ دیگری:
                case "InnerBorderTable":
                    // 🐞 فقط خطوطِ داخلی (بینِ سلول‌ها/ردیف‌ها)، بدونِ بوردرِ
                    // دورِ کلِ جدول.
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
                    span.BorderMode = "inner";
                    span.WidthMode = "equal";
                    break;

                case "FirstRowBorderTable":
                    // 🐞 بوردرِ دورِ کلِ جدول + یک خط زیرِ ردیفِ اول (مثلاً برایِ
                    // جدولی که فقط سرستون از بدنه جدا شود).
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
                    span.BorderMode = "firstRowOuter";
                    span.WidthMode = "equal";
                    break;

                default:
                    // 🐞 استایل‌های ناشناخته (Grid Table*، Table Grid، استایلِ
                    // سفارشیِ CommonTable و هر جدولِ عمومیِ دیگر) به‌صورتِ پیش‌فرض
                    // مثلِ CommonTable وفادار رندر می‌شوند: بوردر/رنگِ هر سلول
                    // عیناً از سند، عرضِ مطلق، و تصمیمِ اسکرول بر پایه‌ی عرضِ کل
                    // نسبت به دستگاه. این همان «یک استایلِ کلی برای همه‌ی
                    // جداول» است. استایل‌های خاصِ اپ (Dotted/Outside/ColumnStack/
                    // MultiColumn/Compact/…) دست‌نخورده و با کاربردِ قبلی می‌مانند.
                    // برای کتاب‌های قدیمی بی‌اثر است (فیلدها فقط در استخراجِ تازه
                    // نوشته می‌شوند و فلاتر آن‌وقت به fallbackِ نام‌محور می‌رود).
                    span.BorderMode = "cell";
                    span.WidthMode = "natural";
                    break;
            }
        }

        // 🐞 هم‌ارزِ سمتِ سی‌شارپِ همان منطقِ isSolidBorderStyle در فلاتر:
        // فقط سبک‌های پیوسته (single/double/…)، نه none/nil/نقطه‌چین/خط‌چین.
        private static bool IsSolidBorderStyle(string val)
        {
            if (string.IsNullOrEmpty(val)) return false;
            var v = val.ToLowerInvariant();
            if (v == "none" || v == "nil") return false;
            if (v.Contains("dot") || v.Contains("dash")) return false;
            return true;
        }

        // آیا خودِ سند برایِ این جدول بوردرِ solid تعریف کرده؟ (یا در سطحِ
        // خودِ جدول، یا در سطحِ هر سلولی — هر کدام کافی است.)
        private static bool HasSolidBorderInSource(SpanData span)
        {
            if (IsSolidBorderStyle(span.Borders?.Val)) return true;
            if (span.TableRows == null) return false;
            foreach (var row in span.TableRows)
            {
                if (row?.Cells == null) continue;
                foreach (var cell in row.Cells)
                {
                    var cb = cell?.Borders;
                    if (cb == null) continue;
                    if (IsSolidBorderStyle(cb.Top?.Val) || IsSolidBorderStyle(cb.Bottom?.Val)
                        || IsSolidBorderStyle(cb.Left?.Val) || IsSolidBorderStyle(cb.Right?.Val))
                        return true;
                }
            }
            return false;
        }
    }
}