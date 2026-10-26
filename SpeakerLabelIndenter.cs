using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CustomLayoutGenerator
{
    /// <summary>
    /// 🌟 A-5: تورفتگیِ آویزانِ برچسبِ گوینده در اسکریپتِ صوتی (دکمه‌ی «Indent Speaker Labels»).
    ///
    /// در ردیفِ دومِ هر جدولِ اسکریپت (همان جدول‌هایی که ProcessAudioScriptWordFile، AudioMarkerReview و
    /// align_engine می‌خوانند: ≥ ۲ ردیف و ردیفِ اول = نامِ ترک) بندی که با برچسب شروع می‌شود
    /// («Man: …»، «Bank clerk: …»؛ معمولاً بولد، بعدش «:» و فاصله) مثلِ بندِ لیستِ شماره‌دار چیده می‌شود:
    ///   - فاصله‌ی بعد از «:» یک Tab می‌شود («Man:⇥Hello»)؛
    ///   - بند تورفتگیِ آویزان می‌گیرد: Left = Hanging = پهنای بلندترین برچسبِ همان جدول + کمی فاصله.
    ///     پس متنِ همه‌ی گوینده‌های یک ترک از یک ستون شروع می‌شود و سطرهای بعدی زیرِ متن می‌آیند،
    ///     نه زیرِ برچسب.
    ///
    /// - جدولی که بندِ لیستِ شماره‌دار/بولت‌دار دارد (A-4) دست نمی‌خورد.
    /// - بندی که برچسب دارد ولی بعدش متنی نیست (برچسبِ تنها در یک بند) دست نمی‌خورد.
    /// - تعدادِ کلمه‌ها عوض نمی‌شود: AudioMarkerReview و align_engine «Tab» را مثلِ فاصله می‌خوانند،
    ///   پس پیش‌نویس‌ها و مارکرهای موجود معتبر می‌مانند. روی فایلِ _markers هم کار می‌کند: مارکرها
    ///   («Man:[1500] Hello») سرِ جایشان می‌مانند و فقط فاصله Tab می‌شود.
    /// - دوباره اجرا کردن بی‌خطر است: Tabِ موجود می‌ماند و تورفتگی دوباره حساب می‌شود.
    /// </summary>
    public static class SpeakerLabelIndenter
    {
        public sealed class Result
        {
            public int Tables;              // جدول‌های اسکریپت
            public int SkippedListTables;   // جدول‌هایی که لیست داشتند و دست نخوردند
            public int IndentedTables;      // جدول‌هایی که دست‌کم یک بندشان تورفتگی گرفت
            public int Paragraphs;          // بندهای برچسب‌دار که تورفتگی گرفتند
            public int TabsInserted;        // فاصله‌هایی که Tab شدند (بقیه از قبل Tab داشتند)
            public int CappedLabels;        // برچسب‌هایی که از سقفِ پهنا بلندتر بودند
            public bool Changed => Paragraphs > 0;
        }

        /// <summary>اطلاعاتِ اندازه‌گیریِ یک برچسب.</summary>
        public readonly record struct LabelText(string Text, string FontName, double SizePt, bool Bold);

        /// <summary>فاصله‌ی بعد از بلندترین برچسب (pt) و سقفِ تورفتگی (pt، ۲٫۵ اینچ).</summary>
        private const double GapPt = 6.0;
        private const double MaxHangPt = 180.0;
        /// <summary>تورفتگی رو به بالا به مضربِ ۰٫۰۵ اینچ (۷۲ twip) گرد می‌شود.</summary>
        private const int RoundTwips = 72;

        private static readonly Regex MarkerRegex = new Regex(
            @"\[(?:(?:\d{1,2}):)?\d{1,2}:\d{2}(?:\.\d+)?\]|\[\d+\]", RegexOptions.Compiled);

        // برچسب: اولِ بند، تا اولین «:» که بعدش فاصله/Tab و بعد متن می‌آید.
        private static readonly Regex LabelHead = new Regex(
            @"^(?<lead>\s*)(?<label>[^\s:][^:\n\t]{0,38}:)(?<ws>[ \t ]+)(?=\S)", RegexOptions.Compiled);

        /// <param name="measurePt">پهنای متنِ برچسب به pt (در برنامه با System.Drawing).</param>
        public static Result Apply(WordprocessingDocument doc, Func<LabelText, double> measurePt)
        {
            var result = new Result();
            var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("No main document part.");
            var body = mainPart.Document?.Body ?? throw new InvalidOperationException("No document body.");
            var fonts = new FontResolver(doc);
            var numberedStyles = NumberedStyleIds(mainPart);

            foreach (var table in body.Descendants<Table>().ToList())
            {
                var rows = table.Elements<TableRow>().ToList();
                if (rows.Count < 2 || string.IsNullOrEmpty(rows[0].InnerText.Trim())) continue;
                result.Tables++;

                var paragraphs = rows[1].Elements<TableCell>().SelectMany(c => c.Elements<Paragraph>()).ToList();
                if (paragraphs.Any(p => IsListParagraph(p, numberedStyles)))
                {
                    result.SkippedListTables++;
                    continue;
                }

                // ۱) بندهای برچسب‌دار و پهنای برچسب‌ها
                var found = new List<(Paragraph Para, Flat Flat, Match M, double WidthPt)>();
                foreach (var p in paragraphs)
                {
                    var flat = Flatten(p);
                    var m = LabelHead.Match(flat.Clean);
                    if (!m.Success) continue;
                    string label = m.Groups["label"].Value;
                    var labelRuns = flat.RunsOf(m.Groups["label"].Index, m.Groups["label"].Length);
                    bool bold = labelRuns.Count > 0 && labelRuns.All(r => IsBold(r, mainPart));
                    char firstChar = label.FirstOrDefault(char.IsLetterOrDigit);
                    bool ok = (AudioScriptTiming.IsLabel(label) && char.IsUpper(firstChar))
                              || (bold && label.Length <= 30 && label.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length <= 4);
                    if (!ok) continue;

                    var firstRun = labelRuns.FirstOrDefault();
                    string font = firstRun != null ? fonts.GetEffectiveFontName(firstRun) ?? "Calibri" : "Calibri";
                    double size = firstRun != null ? FontSizePt(firstRun, p, mainPart) : 11.0;
                    double width = measurePt(new LabelText(label, font, size, bold));
                    found.Add((p, flat, m, width));
                }
                if (found.Count == 0) continue;

                // ۲) تورفتگیِ مشترکِ همین جدول
                double maxW = found.Max(f => f.WidthPt);
                double hangPt = Math.Min(maxW + GapPt, MaxHangPt);
                int hang = (int)Math.Ceiling(hangPt * 20.0 / RoundTwips) * RoundTwips;
                result.CappedLabels += found.Count(f => f.WidthPt + GapPt > MaxHangPt);

                foreach (var (p, flat, m, _) in found)
                {
                    var ws = m.Groups["ws"];
                    if (ws.Value != "\t")
                    {
                        ReplaceWithTab(flat, ws.Index, ws.Length);
                        result.TabsInserted++;
                    }
                    SetHangingIndent(p, hang);
                    result.Paragraphs++;
                }
                result.IndentedTables++;
            }
            return result;
        }

        // ------------------------------------------------------------------

        /// <summary>متنِ بند (Runهای مستقیم، مثلِ FlattenRow1Text) بدونِ مارکرها، با نقشه‌ی هر نویسه.</summary>
        private sealed class Flat
        {
            public string Clean = "";
            // به‌ازای هر نویسه‌ی Clean: Text node و آفستش در آن، یا TabChar (Text = null)
            public List<(Text? Node, int Local, Run Run, OpenXmlElement? Tab)> Map = new();

            /// <summary>Runهای نویسه‌های غیرِفاصله‌ی بازه، به ترتیب.</summary>
            public List<Run> RunsOf(int start, int length)
            {
                var runs = new List<Run>();
                for (int i = start; i < start + length && i < Map.Count; i++)
                    if (!char.IsWhiteSpace(Clean[i]) && !runs.Contains(Map[i].Run)) runs.Add(Map[i].Run);
                return runs;
            }
        }

        private static Flat Flatten(Paragraph p)
        {
            var flat = new Flat();
            var sb = new StringBuilder();
            foreach (var run in p.Elements<Run>())
            {
                foreach (var child in run.ChildElements)
                {
                    if (child is Text t)
                    {
                        string s = t.Text ?? "";
                        // مارکرها حذف، بقیه‌ی نویسه‌ها با جایشان
                        var keep = new bool[s.Length];
                        for (int i = 0; i < s.Length; i++) keep[i] = true;
                        foreach (Match mk in MarkerRegex.Matches(s))
                            for (int i = mk.Index; i < mk.Index + mk.Length; i++) keep[i] = false;
                        for (int i = 0; i < s.Length; i++)
                        {
                            if (!keep[i]) continue;
                            sb.Append(s[i]);
                            flat.Map.Add((t, i, run, null));
                        }
                    }
                    else if (child is TabChar)
                    {
                        sb.Append('\t');
                        flat.Map.Add((null, 0, run, child));
                    }
                    else if (child is Break)
                    {
                        sb.Append('\n');
                        flat.Map.Add((null, 0, run, null));
                    }
                }
            }
            flat.Clean = sb.ToString();
            return flat;
        }

        /// <summary>
        /// فاصله‌های [start, start+length) از Clean حذف و یک Tab جای اولی گذاشته می‌شود — داخلِ همان Run،
        /// تا استایلِ برچسب و متن دست نخورد. مارکری که بینِ «:» و فاصله است («clerk:[500] text») جایش می‌ماند.
        /// </summary>
        private static void ReplaceWithTab(Flat flat, int start, int length)
        {
            var chars = flat.Map.Skip(start).Take(length).ToList();
            var first = chars[0];

            // حذفِ نویسه‌ها از Text nodeها (از آخر به اول تا آفست‌ها جابه‌جا نشوند)
            var byNode = chars.Where(c => c.Node != null).GroupBy(c => c.Node!);
            var tabsToRemove = chars.Where(c => c.Tab != null).Select(c => c.Tab!).ToList();

            if (first.Node != null)
            {
                var node = first.Node;
                int split = first.Local;
                string text = node.Text ?? "";
                // نویسه‌های حذفیِ همین node
                var drop = new HashSet<int>(chars.Where(c => c.Node == node).Select(c => c.Local));
                string before = text.Substring(0, split);
                var after = new StringBuilder();
                for (int i = split; i < text.Length; i++) if (!drop.Contains(i)) after.Append(text[i]);

                node.Text = before;
                node.Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve);
                var tab = new TabChar();
                node.InsertAfterSelf(tab);
                if (after.Length > 0)
                    tab.InsertAfterSelf(new Text(after.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                if (before.Length == 0) node.Remove();
            }
            else
            {
                // اولین نویسه خودش Tab بود («: ⇥» نه؛ یعنی « ⇥»): Tab همان می‌ماند، بقیه حذف
                tabsToRemove.Remove(first.Tab!);
            }

            foreach (var g in byNode)
            {
                if (g.Key == first.Node) continue;
                var drop = new HashSet<int>(g.Select(c => c.Local));
                string text = g.Key.Text ?? "";
                var keep = new StringBuilder();
                for (int i = 0; i < text.Length; i++) if (!drop.Contains(i)) keep.Append(text[i]);
                g.Key.Text = keep.ToString();
                g.Key.Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve);
                if (keep.Length == 0) g.Key.Remove();
            }
            foreach (var t in tabsToRemove) t.Remove();
        }

        private static void SetHangingIndent(Paragraph p, int hangTwips)
        {
            var pPr = p.ParagraphProperties ?? p.PrependChild(new ParagraphProperties());
            var ind = pPr.Indentation;
            // نقطه‌ی شروعِ خطِ اولِ فعلی (اگر قبلاً آویزان شده، همان لبه‌ی برچسب)
            int left = ParseTwips(ind?.Left?.Value) ?? ParseTwips(ind?.Start?.Value) ?? 0;
            int first = ParseTwips(ind?.FirstLine?.Value) ?? -(ParseTwips(ind?.Hanging?.Value) ?? 0);
            int baseTwips = Math.Max(0, left + Math.Min(first, 0));

            if (ind == null)
            {
                ind = new Indentation();
                pPr.Indentation = ind;
            }
            ind.Left = (baseTwips + hangTwips).ToString();
            ind.Hanging = hangTwips.ToString();
            ind.Start = null;
            ind.FirstLine = null;
            ind.FirstLineChars = null;
            ind.HangingChars = null;
            ind.LeftChars = null;
            ind.StartCharacters = null;
        }

        private static int? ParseTwips(string? v) =>
            int.TryParse(v, out int n) ? n : (double.TryParse(v, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double d) ? (int)Math.Round(d) : null);

        private static bool IsBold(Run run, MainDocumentPart mainPart)
        {
            var b = run.RunProperties?.Bold;
            if (b != null) return b.Val == null || b.Val.Value;
            var styles = mainPart.StyleDefinitionsPart?.Styles;
            string? id = run.RunProperties?.RunStyle?.Val?.Value;
            for (int guard = 0; id != null && styles != null && guard < 10; guard++)
            {
                var st = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == id);
                var sb = st?.StyleRunProperties?.Bold;
                if (sb != null) return sb.Val == null || sb.Val.Value;
                id = st?.BasedOn?.Val?.Value;
            }
            return false;
        }

        /// <summary>اندازه‌ی فونت (pt): Run ← استایلِ کاراکتری ← استایلِ بند ← docDefaults ← ۱۱.</summary>
        private static double FontSizePt(Run run, Paragraph p, MainDocumentPart mainPart)
        {
            string? sz = run.RunProperties?.FontSize?.Val?.Value
                         ?? StyleSize(mainPart, run.RunProperties?.RunStyle?.Val?.Value)
                         ?? StyleSize(mainPart, p.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
                         ?? mainPart.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault
                                ?.RunPropertiesBaseStyle?.FontSize?.Val?.Value;
            return double.TryParse(sz, out double half) ? half / 2.0 : 11.0;
        }

        private static string? StyleSize(MainDocumentPart mainPart, string? styleId)
        {
            var styles = mainPart.StyleDefinitionsPart?.Styles;
            for (int guard = 0; styleId != null && styles != null && guard < 10; guard++)
            {
                var st = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
                var v = st?.StyleRunProperties?.FontSize?.Val?.Value;
                if (!string.IsNullOrEmpty(v)) return v;
                styleId = st?.BasedOn?.Val?.Value;
            }
            return null;
        }

        /// <summary>مثلِ A-4: استایل‌های بندی که خودشان numPr با numId غیرِ صفر دارند.</summary>
        private static HashSet<string> NumberedStyleIds(MainDocumentPart mainPart)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var styles = mainPart.StyleDefinitionsPart?.Styles;
            if (styles == null) return ids;
            foreach (var st in styles.Elements<Style>())
            {
                int? num = st.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
                if (num != null && num != 0 && st.StyleId?.Value != null) ids.Add(st.StyleId.Value);
            }
            return ids;
        }

        private static bool IsListParagraph(Paragraph p, HashSet<string> numberedStyles)
        {
            int? num = p.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            if (num != null) return num != 0;
            string? style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            return style != null && numberedStyles.Contains(style);
        }
    }
}
