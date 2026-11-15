using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CustomLayoutGenerator; // نام namespace را با پروژه‌ی خودتان یکی کنید

public enum JoinMode
{
    Strict,  // پاراگراف دوم باید با حرف کوچک لاتین شروع شود
    Relaxed  // فقط «پاراگراف اول بدون علامت پایان جمله» کافی است
}

public sealed class JoinItem
{
    public int? Page { get; init; }
    public string EndOfFirst { get; init; } = "";
    public string StartOfNext { get; init; } = "";
}

public sealed class JoinReport
{
    public List<JoinItem> Items { get; } = new();
    public string? ReportPath { get; set; }

    public string ToSummary()
    {
        string s = $"\n\nپاراگراف‌های شکسته: {Items.Count} مورد ادغام شد";
        if (ReportPath != null) s += $"\nگزارش:\n{ReportPath}";
        return s;
    }
}

public static class ParagraphJoiner
{
    const string Terminators = ".!?…:;؟؛。";
    const string ClosingChars = "\"'”’»)]}›";
    const string OpeningChars = "\"'“‘«([{‹";

    /// <summary>
    /// از فایل ورودی یک کپی می‌سازد و پاراگراف‌های شکسته را در آن ادغام می‌کند.
    /// protectedTexts: متن‌هایی (مثل نام گوینده‌ها) که هرگز نباید ادغام شوند.
    /// excludedStyleNames: نام استایل‌هایی که پاراگراف‌هایشان نباید ادغام شوند (مثلاً NotSpoken).
    /// </summary>
    public static JoinReport JoinFile(
        string inputPath, string outputPath, JoinMode mode,
        IEnumerable<string> protectedTexts, IEnumerable<string> excludedStyleNames)
    {
        var protectedSet = new HashSet<string>(
            protectedTexts.Select(Normalize).Where(s => s.Length > 0),
            StringComparer.OrdinalIgnoreCase);
        var excludedNames = excludedStyleNames.ToList();

        File.Copy(inputPath, outputPath, overwrite: true);
        using var doc = WordprocessingDocument.Open(outputPath, true);
        var main = doc.MainDocumentPart!;
        var body = main.Document!.Body!;

        // شناسه‌ی استایل‌های مستثنا (از روی نام یا شناسه)
        var excludedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var styles = main.StyleDefinitionsPart?.Styles;
        if (styles != null)
        {
            foreach (var s in styles.Elements<Style>())
            {
                if (s.Type?.Value != StyleValues.Paragraph || s.StyleId?.Value is not { } id) continue;
                if (excludedNames.Any(n =>
                        string.Equals(n, s.StyleName?.Val?.Value, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(n, id, StringComparison.OrdinalIgnoreCase)))
                    excludedIds.Add(id);
            }
        }

        // شماره‌ی صفحه‌ی تقریبی هر پاراگراف (قبل از هر تغییر)؛ از LastRenderedPageBreak که Word هنگام ذخیره می‌گذارد
        var pageOf = new Dictionary<Paragraph, int>();
        int page = 1;
        bool anyBreak = false;
        foreach (var el in body.Descendants())
        {
            if (el is LastRenderedPageBreak) { page++; anyBreak = true; }
            else if (el is Paragraph pp) pageOf[pp] = page;
        }
        if (!anyBreak) pageOf.Clear(); // فایل هیچ‌وقت در Word ذخیره نشده؛ صفحه نامعلوم است

        var report = new JoinReport();

        // لیست را از قبل می‌گیریم چون حین کار پاراگراف حذف می‌کنیم
        foreach (var p in body.Descendants<Paragraph>().ToList())
        {
            if (p.Parent == null) continue; // قبلاً در یک ادغام حذف شده

            // زنجیره‌ای: تا وقتی بعدی قابل ادغام است ادامه می‌دهیم
            while (p.NextSibling() is Paragraph next &&
                   CanJoin(p, next, mode, protectedSet, excludedIds))
            {
                int? nextPage = pageOf.TryGetValue(next, out var pg) ? pg : null;
                var item = new JoinItem
                {
                    Page = nextPage,
                    EndOfFirst = Tail(p.InnerText),
                    StartOfNext = Head(next.InnerText)
                };

                Join(p, next);
                report.Items.Add(item);
            }
        }

        main.Document.Save();

        // نوشتن گزارش کنار فایل خروجی
        string reportPath = Path.Combine(
            Path.GetDirectoryName(outputPath)!,
            Path.GetFileNameWithoutExtension(outputPath) + "_joined_report.txt");
        WriteReport(report, reportPath, Path.GetFileName(inputPath), mode);

        return report;
    }

    // ───────────────────────── تشخیص ─────────────────────────

    static bool CanJoin(Paragraph a, Paragraph b, JoinMode mode,
                        HashSet<string> protectedSet, HashSet<string> excludedIds)
    {
        string ta = a.InnerText.TrimEnd();
        string tb = b.InnerText.TrimStart();
        if (ta.Length == 0 || tb.Length == 0) return false;

        // استایل پاراگرافی باید یکسان باشد
        string? sa = StyleId(a), sb = StyleId(b);
        if (sa != sb) return false;
        if (sa != null && (excludedIds.Contains(sa) || IsHeadingStyle(sa))) return false;

        // فهرست، سرفصل، شکست صفحه/ستون/بخش
        if (HasBlockers(a) || HasBlockers(b)) return false;

        // متن‌های محافظت‌شده (نام گوینده‌ها)
        if (protectedSet.Contains(Normalize(ta)) || protectedSet.Contains(Normalize(tb)))
            return false;

        // پاراگراف اول نباید جمله را تمام کرده باشد
        if (EndsSentence(ta)) return false;

        // پاراگراف دوم باید با حرف شروع شود (نه عدد یا علامت)
        char first = FirstLetterChar(tb);
        if (first == '\0') return false;

        return mode == JoinMode.Relaxed || char.IsLower(first);
    }

    static string? StyleId(Paragraph p) => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

    static bool IsHeadingStyle(string styleId) =>
        new[] { "Heading", "Title", "Subtitle", "TOC", "Caption" }
            .Any(x => styleId.StartsWith(x, StringComparison.OrdinalIgnoreCase));

    static bool HasBlockers(Paragraph p)
    {
        var pPr = p.ParagraphProperties;
        if (pPr?.NumberingProperties != null) return true; // آیتم فهرست
        if (pPr?.OutlineLevel != null) return true;        // سرفصل
        if (pPr?.SectionProperties != null) return true;   // شکست بخش
        return p.Descendants<Break>().Any(br =>
            br.Type?.Value == BreakValues.Page || br.Type?.Value == BreakValues.Column);
    }

    static bool EndsSentence(string s)
    {
        int i = s.Length - 1;
        while (i >= 0 && ClosingChars.IndexOf(s[i]) >= 0) i--;
        return i >= 0 && Terminators.IndexOf(s[i]) >= 0;
    }

    // اولین کاراکتر متن بعد از گیومه/پرانتز باز؛ اگر حرف نبود '\0'
    static char FirstLetterChar(string s)
    {
        foreach (char c in s)
        {
            if (OpeningChars.IndexOf(c) >= 0) continue;
            return char.IsLetter(c) ? c : '\0';
        }
        return '\0';
    }

    // ───────────────────────── ادغام ─────────────────────────

    static void Join(Paragraph a, Paragraph b)
    {
        // ۱) آخرین Text غیرخالی پاراگراف اول: فاصله‌های انتها را برمی‌داریم
        var textsA = a.Descendants<Text>().ToList();
        Text? last = null;
        for (int i = textsA.Count - 1; i >= 0; i--)
        {
            if (string.IsNullOrWhiteSpace(textsA[i].Text)) { textsA[i].Text = ""; continue; }
            last = textsA[i];
            break;
        }

        if (last != null)
        {
            string t = last.Text.TrimEnd();
            bool softHyphen = t.EndsWith('\u00AD');
            bool hyphen = t.EndsWith('-');

            if (softHyphen) t = t.TrimEnd('\u00AD'); // خط‌تیره‌ی نرم: کلمه به هم می‌چسبد
            last.Text = (softHyphen || hyphen) ? t : t + " ";
            last.Space = SpaceProcessingModeValues.Preserve;
        }

        // ۲) فاصله‌های ابتدای پاراگراف دوم را برمی‌داریم
        foreach (var t in b.Descendants<Text>())
        {
            if (string.IsNullOrWhiteSpace(t.Text)) { t.Text = ""; continue; }
            t.Text = t.Text.TrimStart();
            break;
        }

        // ۳) انتقال محتوای پاراگراف دوم (بدون تغییر فرمت) و حذف آن
        foreach (var child in b.ChildElements.Where(c => c is not ParagraphProperties).ToList())
        {
            child.Remove();
            a.Append(child);
        }
        b.Remove();
    }

    // ───────────────────────── گزارش ─────────────────────────

    static string Tail(string s)
    {
        s = s.TrimEnd();
        return s.Length <= 40 ? s : "…" + s.Substring(s.Length - 40);
    }

    static string Head(string s)
    {
        s = s.TrimStart();
        return s.Length <= 40 ? s : s.Substring(0, 40) + "…";
    }

    static string Normalize(string s)
    {
        s = s.Replace("\u200E", "").Replace("\u200F", "")
             .Replace('ي', 'ی').Replace('ك', 'ک');
        return Regex.Replace(s, @"\s+", " ").Trim();
    }

    static void WriteReport(JoinReport r, string path, string sourceName, JoinMode mode)
    {
        if (r.Items.Count == 0) return;

        var sb = new StringBuilder();
        sb.AppendLine($"گزارش ادغام پاراگراف‌های شکسته — {sourceName}");
        sb.AppendLine($"حالت: {(mode == JoinMode.Strict ? "سخت‌گیرانه" : "آزاد")} | تعداد ادغام: {r.Items.Count}");
        sb.AppendLine("شماره‌ی صفحه تقریبی و مربوط به فایل اصلی است. محل ادغام بین دو متن داخل « » است.");
        sb.AppendLine();

        foreach (var it in r.Items)
        {
            sb.AppendLine($"[ص {(it.Page?.ToString() ?? "؟")}]");
            sb.AppendLine($"    «{it.EndOfFirst}»");
            sb.AppendLine($"    «{it.StartOfNext}»");
        }

        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        r.ReportPath = path;
    }
}