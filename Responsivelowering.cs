using System.Collections.Generic;

namespace WordToJsonParser
{
    /// <summary>
    /// نام‌های استایلِ جدولِ Word را به primitiveهای declarativeِ ریسپانسیو «تبدیل»
    /// می‌کند تا فلاتر به‌جای switch روی نام استایل، فقط با چند فیلدِ عمومی کار کند.
    /// افزودن یک استایلِ جدید = یک case در این switch (بدون تغییر در فلاتر).
    /// روی spanها به‌صورت بازگشتی کار می‌کند (جدولِ تودرتو، innerSpans و سلول‌ها).
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
                    break;

                case "BorderedTable":
                    span.ResponsiveStrategy = "horizontalScroll";
                    span.Borders = span.Borders ?? new BorderDetail();
                    if (string.IsNullOrEmpty(span.Borders.Val)) span.Borders.Val = "single";
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
                    break;

                default:
                    // جدولِ داده‌ی بی‌نام: امن‌ترین پیش‌فرض
                    span.ResponsiveStrategy = "horizontalScroll";
                    break;
            }
        }
    }
}