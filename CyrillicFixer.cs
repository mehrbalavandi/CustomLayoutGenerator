using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CustomLayoutGenerator; // نام namespace را با پروژه‌ی خودتان یکی کنید

// یک مورد در گزارش (اصلاح‌شده یا ردشده)
public sealed class CyrillicItem
{
    public string Part { get; init; } = "";
    public int? Page { get; init; }          // null یعنی صفحه نامعلوم است
    public string Word { get; init; } = "";  // کلمه‌ی اصلی (قبل از اصلاح)
    public string? Fixed { get; init; }      // کلمه‌ی اصلاح‌شده؛ null یعنی رد شده
    public string Context { get; init; } = "";
}

public sealed class CyrillicReport
{
    public bool Enabled { get; set; }
    public List<CyrillicItem> Changes { get; } = new();
    public List<CyrillicItem> Skipped { get; } = new();
    public int ReplacedChars { get; set; }
    public string? ReportPath { get; set; }

    // خلاصه‌ی فارسی برای نمایش در MessageBox
    public string ToSummary()
    {
        if (!Enabled) return "";

        string s = $"\n\nسیریلیک: {Changes.Count} کلمه اصلاح شد ({ReplacedChars} کاراکتر)";
        if (Skipped.Count > 0)
            s += $"، {Skipped.Count} کلمه‌ی سیریلیک واقعی دست نخورد";
        if (ReportPath != null)
            s += $"\nگزارش:\n{ReportPath}";
        return s;
    }
}

public static class CyrillicFixer
{
    // حروف سیریلیکی که شکل لاتین یکسان دارند.
    // اگر مورد دیگری لازم بود، همین‌جا اضافه کنید.
    static readonly Dictionary<char, char> Map = new()
    {
        // بزرگ
        ['\u0410'] = 'A', // А
        ['\u0412'] = 'B', // В
        ['\u0415'] = 'E', // Е
        ['\u041A'] = 'K', // К
        ['\u041C'] = 'M', // М
        ['\u041D'] = 'H', // Н
        ['\u041E'] = 'O', // О
        ['\u0420'] = 'P', // Р
        ['\u0421'] = 'C', // С
        ['\u0422'] = 'T', // Т
        ['\u0423'] = 'Y', // У
        ['\u0425'] = 'X', // Х
        ['\u0406'] = 'I', // І
        ['\u0408'] = 'J', // Ј
        ['\u0405'] = 'S', // Ѕ
        ['\u04AE'] = 'Y', // Ү
        ['\u04BA'] = 'H', // Һ
        ['\u051A'] = 'Q', // Ԛ
        ['\u051C'] = 'W', // Ԝ

        // کوچک
        ['\u0430'] = 'a', // а
        ['\u0441'] = 'c', // с
        ['\u0435'] = 'e', // е
        ['\u043E'] = 'o', // о
        ['\u0440'] = 'p', // р
        ['\u0443'] = 'y', // у
        ['\u0445'] = 'x', // х
        ['\u0456'] = 'i', // і
        ['\u0458'] = 'j', // ј
        ['\u0455'] = 's', // ѕ
        ['\u04BB'] = 'h', // һ
        ['\u0501'] = 'd', // ԁ
        ['\u051B'] = 'q', // ԛ
        ['\u051D'] = 'w', // ԝ
    };

    static bool IsCyrillic(char c) =>
        (c >= '\u0400' && c <= '\u052F') ||
        (c >= '\u1C80' && c <= '\u1C8F') ||
        (c >= '\u2DE0' && c <= '\u2DFF') ||
        (c >= '\uA640' && c <= '\uA69F');

    // کاراکتر «کلمه»: حرف، رقم یا علامت ترکیبی
    static bool IsWordChar(char c) =>
        char.IsLetterOrDigit(c) ||
        CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark;

    /// <summary>
    /// حروف سیریلیک هم‌شکل را در متن اصلی، سربرگ، فوتر، Footnote و Endnote اصلاح می‌کند.
    /// </summary>
    public static CyrillicReport Fix(WordprocessingDocument doc)
    {
        var report = new CyrillicReport { Enabled = true };
        var main = doc.MainDocumentPart!;

        if (main.Document != null)
            FixRoot(main.Document, "متن اصلی", trackPages: true, report);

        foreach (var h in main.HeaderParts)
            if (h.Header != null) FixRoot(h.Header, "سربرگ", false, report);

        foreach (var f in main.FooterParts)
            if (f.Footer != null) FixRoot(f.Footer, "فوتر", false, report);

        if (main.FootnotesPart?.Footnotes is { } footnotes)
            FixRoot(footnotes, "Footnote", false, report);

        if (main.EndnotesPart?.Endnotes is { } endnotes)
            FixRoot(endnotes, "Endnote", false, report);

        return report;
    }

    static void FixRoot(OpenXmlPartRootElement root, string partName, bool trackPages, CyrillicReport report)
    {
        // شماره‌ی صفحه‌ی تقریبی هر Text؛ فقط از نشانه‌های LastRenderedPageBreak که Word هنگام ذخیره می‌گذارد
        Dictionary<Text, int>? pages = null;
        if (trackPages)
        {
            pages = new Dictionary<Text, int>();
            int page = 1;
            bool anyBreak = false;
            foreach (var el in root.Descendants())
            {
                if (el is LastRenderedPageBreak) { page++; anyBreak = true; }
                else if (el is Text t) pages[t] = page;
            }
            if (!anyBreak) pages = null; // فایل هیچ‌وقت در Word ذخیره نشده؛ صفحه نامعلوم است
        }

        bool changed = false;
        foreach (var p in root.Descendants<Paragraph>().ToList())
            changed |= FixParagraph(p, partName, pages, report);

        if (changed) root.Save();
    }

    static bool FixParagraph(Paragraph p, string partName, Dictionary<Text, int>? pages, CyrillicReport report)
    {
        // فقط Textهای خودِ همین پاراگراف (نه پاراگراف‌های تو در تو)
        var texts = p.Descendants<Text>()
            .Where(t => ReferenceEquals(t.Ancestors<Paragraph>().FirstOrDefault(), p))
            .ToList();
        if (texts.Count == 0) return false;

        // متن کل پاراگراف را یکی می‌کنیم تا کلمه‌ی شکسته بین چند Run هم درست تشخیص داده شود
        var starts = new int[texts.Count];
        var lens = new int[texts.Count];
        var sb = new StringBuilder();
        for (int n = 0; n < texts.Count; n++)
        {
            string s = texts[n].Text ?? "";
            starts[n] = sb.Length;
            lens[n] = s.Length;
            sb.Append(s);
        }

        string full = sb.ToString();
        if (!full.Any(IsCyrillic)) return false; // مسیر سریع

        // شماره‌ی صفحه‌ی Textی که این موقعیت داخل آن است
        int? PageAt(int pos)
        {
            if (pages == null) return null;
            for (int n = 0; n < texts.Count; n++)
                if (pos >= starts[n] && pos < starts[n] + lens[n])
                    return pages.TryGetValue(texts[n], out var pg) ? pg : null;
            return null;
        }

        var chars = full.ToCharArray(); // جایگزینی یک‌به‌یک است، پس موقعیت‌ها ثابت می‌مانند
        bool modified = false;
        int i = 0;

        while (i < chars.Length)
        {
            if (!IsWordChar(chars[i])) { i++; continue; }

            int j = i;
            while (j < chars.Length && IsWordChar(chars[j])) j++;

            // بررسی کلمه‌ی [i, j)
            bool hasCyrillic = false, hasUnmapped = false;
            for (int k = i; k < j; k++)
            {
                if (!IsCyrillic(chars[k])) continue;
                hasCyrillic = true;
                if (!Map.ContainsKey(chars[k])) hasUnmapped = true;
            }

            if (hasCyrillic)
            {
                string word = full.Substring(i, j - i);
                string context = GetContext(full, i, j);
                int? page = PageAt(i);

                if (hasUnmapped)
                {
                    // سیریلیک واقعی (مثلاً کلمه‌ی روسی)؛ دست نمی‌زنیم
                    report.Skipped.Add(new CyrillicItem
                    {
                        Part = partName,
                        Page = page,
                        Word = word,
                        Fixed = null,
                        Context = context
                    });
                }
                else
                {
                    for (int k = i; k < j; k++)
                    {
                        if (Map.TryGetValue(chars[k], out char latin))
                        {
                            chars[k] = latin;
                            report.ReplacedChars++;
                        }
                    }

                    report.Changes.Add(new CyrillicItem
                    {
                        Part = partName,
                        Page = page,
                        Word = word,
                        Fixed = new string(chars, i, j - i),
                        Context = context
                    });
                    modified = true;
                }
            }

            i = j;
        }

        if (!modified) return false;

        // نوشتن متن اصلاح‌شده به Textهای اصلی (بدون دست زدن به فرمت)
        for (int n = 0; n < texts.Count; n++)
        {
            string seg = new string(chars, starts[n], lens[n]);
            if (seg != texts[n].Text) texts[n].Text = seg;
        }
        return true;
    }

    // کمی متن اطراف کلمه برای پیدا کردن آن با Find در Word
    static string GetContext(string full, int start, int end)
    {
        int from = Math.Max(0, start - 25);
        int to = Math.Min(full.Length, end + 25);
        string c = full.Substring(from, to - from).Replace('\n', ' ').Replace('\t', ' ');
        return (from > 0 ? "…" : "") + c + (to < full.Length ? "…" : "");
    }

    // نمایش کلمه با مشخص کردن حرف سیریلیک و کد یونیکد آن، مثلاً Exam[і U+0456]ner
    static string Mark(string word)
    {
        var sb = new StringBuilder();
        foreach (char c in word)
        {
            if (IsCyrillic(c)) sb.Append($"[{c} U+{(int)c:X4}]");
            else sb.Append(c);
        }
        return sb.ToString();
    }

    /// <summary>
    /// گزارش را در فایل متنی (UTF-8) می‌نویسد؛ اگر موردی نباشد فایلی ساخته نمی‌شود.
    /// </summary>
    public static void WriteReport(CyrillicReport r, string reportPath, string sourceFileName)
    {
        if (r.Changes.Count == 0 && r.Skipped.Count == 0) return;

        var sb = new StringBuilder();
        sb.AppendLine($"گزارش اصلاح حروف سیریلیک — {sourceFileName}");
        sb.AppendLine($"کلمه‌ی اصلاح‌شده: {r.Changes.Count} | کاراکتر: {r.ReplacedChars} | کلمه‌ی ردشده: {r.Skipped.Count}");
        sb.AppendLine("شماره‌ی صفحه تقریبی است (از نشانه‌های ذخیره‌ی Word). متن داخل « » را در فایل اصلی با Find جست‌وجو کنید.");
        sb.AppendLine();

        if (r.Changes.Count > 0)
        {
            sb.AppendLine("=== اصلاح‌شده‌ها ===");
            foreach (var it in r.Changes)
            {
                sb.AppendLine($"[ص {(it.Page?.ToString() ?? "؟")}] ({it.Part}) {Mark(it.Word)}  →  {it.Fixed}");
                sb.AppendLine($"    «{it.Context}»");
            }
            sb.AppendLine();
        }

        if (r.Skipped.Count > 0)
        {
            sb.AppendLine("=== ردشده‌ها (سیریلیک واقعی؛ دست نخورد) ===");
            foreach (var it in r.Skipped)
            {
                sb.AppendLine($"[ص {(it.Page?.ToString() ?? "؟")}] ({it.Part}) {it.Word}");
                sb.AppendLine($"    «{it.Context}»");
            }
        }

        File.WriteAllText(reportPath, sb.ToString(), new UTF8Encoding(true));
        r.ReportPath = reportPath;
    }
}