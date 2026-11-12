using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace YourApp; // نام namespace را با پروژه‌ی خودتان یکی کنید

public static class SpeakerMerger
{
    // کاراکترهایی که اگر پاراگراف به آن‌ها ختم شود، ادغام انجام نمی‌شود
    static readonly char[] Terminators = { '.', ':' };

    // ترتیب استاندارد المان‌ها داخل rPr (ترتیب غلط ممکن است Word را به خطا بیندازد)
    static readonly string[] RPrOrder =
    {
        "rStyle","rFonts","b","bCs","i","iCs","caps","smallCaps","strike","dstrike",
        "outline","shadow","emboss","imprint","noProof","snapToGrid","vanish","webHidden",
        "color","spacing","w","kern","position","sz","szCs","highlight","u","effect",
        "bdr","shd","fitText","vertAlign","rtl","cs","em","lang","eastAsianLayout",
        "specVanish","oMath"
    };

    /// <summary>
    /// از فایل ورودی یک کپی می‌سازد، پاراگراف‌ها را ادغام می‌کند و تعداد ادغام‌ها را برمی‌گرداند.
    /// </summary>
    public static int MergeFile(string inputPath, string outputPath, string styleName)
    {
        File.Copy(inputPath, outputPath, overwrite: true);
        using var doc = WordprocessingDocument.Open(outputPath, true);
        return Process(doc, styleName);
    }

    static int Process(WordprocessingDocument doc, string styleName)
    {
        var main = doc.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart?.Styles
            ?? throw new Exception("فایل بخش استایل ندارد.");

        // پیدا کردن استایل پاراگرافی گوینده از روی نام (یا شناسه)
        var speakerStyle = styles.Elements<Style>().FirstOrDefault(s =>
            s.Type?.Value == StyleValues.Paragraph &&
            (string.Equals(s.StyleName?.Val?.Value, styleName, StringComparison.OrdinalIgnoreCase) ||
             string.Equals(s.StyleId?.Value, styleName, StringComparison.OrdinalIgnoreCase)))
            ?? throw new Exception($"استایل پاراگرافی با نام «{styleName}» در این فایل پیدا نشد.");

        string styleId = speakerStyle.StyleId!.Value!;

        // استایل کاراکتری معادل NotSpoken (موجود را پیدا می‌کند یا می‌سازد)
        string charStyleId = EnsureCharacterStyle(styles, speakerStyle, styleName);

        // لیست را از قبل می‌گیریم چون حین کار پاراگراف حذف می‌کنیم
        var paragraphs = main.Document!.Body!.Descendants<Paragraph>().ToList();
        int merged = 0;

        foreach (var p in paragraphs)
        {
            if (p.Parent == null) continue;                  // قبلاً حذف شده
            if (!HasStyle(p, styleId)) continue;

            string text = p.InnerText.TrimEnd();
            if (text.Length == 0) continue;                  // پاراگراف خالی
            if (Terminators.Contains(text[^1])) continue;    // با . یا : تمام شده

            // پاراگراف بعدی باید واقعاً پاراگراف باشد (نه جدول و ...)
            if (p.NextSibling() is not Paragraph next) continue;
            if (HasStyle(next, styleId)) continue;           // بعدی هم گوینده است
            if (next.InnerText.Trim().Length == 0) continue; // بعدی خالی است

            MergeInto(p, next, charStyleId);
            merged++;
        }

        styles.Save();
        main.Document.Save();
        return merged;
    }

    static bool HasStyle(Paragraph p, string styleId) =>
        p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == styleId;

    static void MergeInto(Paragraph speaker, Paragraph next, string charStyleId)
    {
        // ۱) فاصله‌های انتهای متن گوینده را پاک می‌کنیم و آخرین Text را پیدا می‌کنیم
        var texts = speaker.Descendants<Text>().ToList();
        Text? lastText = null;
        for (int i = texts.Count - 1; i >= 0; i--)
        {
            if (string.IsNullOrWhiteSpace(texts[i].Text)) { texts[i].Text = ""; continue; }
            texts[i].Text = texts[i].Text.TrimEnd();
            texts[i].Space = SpaceProcessingModeValues.Preserve;
            lastText = texts[i];
            break;
        }

        // ۲) استایل کاراکتری NotSpoken را روی همه‌ی Runهای گوینده می‌گذاریم
        foreach (var run in speaker.Descendants<Run>())
            SetRunStyle(run, charStyleId);

        // ۳) اضافه کردن ": " با همان ظاهر گوینده
        var colonRun = new Run();
        var lastRun = lastText?.Ancestors<Run>().FirstOrDefault()
                      ?? speaker.Descendants<Run>().LastOrDefault();
        if (lastRun?.RunProperties != null)
            colonRun.Append(lastRun.RunProperties.CloneNode(true));
        SetRunStyle(colonRun, charStyleId);
        colonRun.Append(new Text(": ") { Space = SpaceProcessingModeValues.Preserve });
        speaker.Append(colonRun);

        // ۴) تنظیمات پاراگراف (از جمله استایل Normal یا هر استایل دیگر) از پاراگراف دوم گرفته می‌شود
        //    تا استایل پاراگرافی NotSpoken روی کل پاراگراف ادغام‌شده نماند
        speaker.ParagraphProperties?.Remove();
        if (next.ParagraphProperties != null)
            speaker.PrependChild((ParagraphProperties)next.ParagraphProperties.CloneNode(true));

        // ۵) انتقال محتوای پاراگراف دوم (بدون هیچ تغییر فرمتی) و حذف آن
        foreach (var child in next.ChildElements.Where(c => c is not ParagraphProperties).ToList())
        {
            child.Remove();
            speaker.Append(child);
        }
        next.Remove();
    }

    // گذاشتن RunStyle روی یک Run (rStyle باید اولین فرزند rPr باشد)
    static void SetRunStyle(Run run, string charStyleId)
    {
        var rp = run.RunProperties ?? run.PrependChild(new RunProperties());
        rp.RunStyle?.Remove();
        rp.PrependChild(new RunStyle { Val = charStyleId });
    }

    // پیدا کردن یا ساختن استایل کاراکتری معادل استایل پاراگرافی گوینده
    static string EnsureCharacterStyle(Styles styles, Style speakerStyle, string styleName)
    {
        // اگر NotSpoken یک Linked Style است، کاراکتر لینک‌شده‌اش را استفاده می‌کنیم
        string? linkId = speakerStyle.LinkedStyle?.Val?.Value;
        if (linkId != null && styles.Elements<Style>().Any(s =>
                s.StyleId?.Value == linkId && s.Type?.Value == StyleValues.Character))
            return linkId;

        // اگر قبلاً با همین روش ساخته شده، دوباره استفاده می‌شود
        string newId = speakerStyle.StyleId!.Value! + "Char";
        if (styles.Elements<Style>().Any(s =>
                s.StyleId?.Value == newId && s.Type?.Value == StyleValues.Character))
            return newId;

        // ساختن استایل کاراکتری جدید با ظاهر کاراکتری NotSpoken
        var charStyle = new Style
        {
            Type = StyleValues.Character,
            StyleId = newId,
            CustomStyle = true
        };
        charStyle.Append(new StyleName { Val = (speakerStyle.StyleName?.Val?.Value ?? styleName) + " Char" });
        if (styles.Elements<Style>().Any(s => s.StyleId?.Value == "DefaultParagraphFont"))
            charStyle.Append(new BasedOn { Val = "DefaultParagraphFont" });
        charStyle.Append(new UIPriority { Val = 99 });

        var rPr = new StyleRunProperties();
        foreach (var s in BuildStyleChain(styles, speakerStyle))
        {
            // استایل پیش‌فرض پاراگراف (Normal) را کپی نمی‌کنیم؛ پاراگراف نهایی خودش آن را دارد
            if (s.Default?.Value == true) continue;
            if (s.StyleRunProperties != null)
                Overlay(rPr, s.StyleRunProperties);
        }
        SortRunProps(rPr);
        charStyle.Append(rPr);

        styles.Append(charStyle);
        return newId;
    }

    // زنجیره‌ی BasedOn یک استایل (پایه اول، مشتق‌شده آخر)
    static List<Style> BuildStyleChain(Styles styles, Style style)
    {
        var chain = new List<Style>();
        Style? cur = style;
        while (cur != null && chain.Count < 20)
        {
            chain.Insert(0, cur);
            string? baseId = cur.BasedOn?.Val?.Value;
            cur = baseId == null ? null
                : styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == baseId);
        }
        return chain;
    }

    // کپی کردن المان‌های source روی target؛ المان هم‌نوع قبلی جایگزین می‌شود
    static void Overlay(OpenXmlElement target, OpenXmlElement source)
    {
        foreach (var child in source.ChildElements.ToList())
        {
            target.ChildElements.FirstOrDefault(e => e.GetType() == child.GetType())?.Remove();
            target.AppendChild(child.CloneNode(true));
        }
    }

    // مرتب‌سازی فرزندان rPr طبق ترتیب اسکیمای Word
    static void SortRunProps(OpenXmlElement rp)
    {
        var sorted = rp.ChildElements
            .OrderBy(e => { int i = Array.IndexOf(RPrOrder, e.LocalName); return i < 0 ? int.MaxValue : i; })
            .Select(e => e.CloneNode(true))
            .ToList();
        rp.RemoveAllChildren();
        foreach (var e in sorted) rp.AppendChild(e);
    }
}