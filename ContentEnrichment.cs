using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CustomLayoutGenerator
{
    /// <summary>
    /// 🌟 غنی‌سازی (ترجمه‌ی فارسی/عربی + کلمات و اصطلاحاتِ تعاملی) — قانون‌های G-1 تا G-7
    /// در RULES-extraction.md.
    ///
    /// ترجمه‌ها دیگر داخلِ JSONِ صفحه‌ها «ساخته» نمی‌شوند؛ جدا از کتاب نگه داشته می‌شوند و در
    /// هر استخراج، همین‌جا به پاراگراف‌ها وصل می‌شوند:
    ///
    ///   {پوشه‌ی کتاب}/enrichment/translations.json   حافظه‌ی ترجمه‌ی همین کتاب
    ///   {پوشه‌ی داده}/_lexicon/lexicon.json            دیکشنریِ مشترکِ همه‌ی کتاب‌ها
    ///   {پوشه‌ی کتاب}/enrichment/pending.json          خروجی: پاراگراف‌هایی که هنوز ترجمه ندارند
    ///
    /// کلیدِ هر پاراگراف «متنِ نرمال‌شده»ی آن است (Normalize)، نه جایش در کتاب. پس تغییرِ
    /// استایل، رنگ، چیدمان، صفحه یا ترتیبِ پاراگراف‌ها ترجمه را از بین نمی‌برد؛ فقط تغییرِ خودِ
    /// متن، آن پاراگراف را به pending.json می‌فرستد.
    /// </summary>
    public static class ContentEnrichment
    {
        public const string EnrichmentFolder = "enrichment";
        public const string TranslationsFile = "translations.json";
        public const string PendingFile = "pending.json";
        public const string LexiconFolder = "_lexicon";
        public const string LexiconFile = "lexicon.json";

        public class Result
        {
            /// <summary>فهرستِ Interactivesِ سطحِ کتاب برای index.json (null اگر دیکشنری خالی است).</summary>
            public List<JObject> Interactives;
            public int TotalParagraphs;
            public int Translated;
            public int Pending;
            public int InteractiveEntries;
            public string PendingPath;
            public string Summary =>
                $"ترجمه: {Translated} از {TotalParagraphs} پاراگراف" +
                (Pending > 0 ? $"؛ بدونِ ترجمه: {Pending} (enrichment/{PendingFile})" : "") +
                $"؛ کلمات و اصطلاحاتِ تعاملی: {InteractiveEntries}";
        }

        private class TmEntry
        {
            public string Fa;
            public string Ar;
            public List<string> Terms = new List<string>();
        }

        // ───────────────────────── کلید ─────────────────────────

        /// <summary>
        /// G-1: کلیدِ پاراگراف. علامت‌های {blk}، نوعِ گیومه و خط‌تیره، فاصله‌های اضافه و
        /// بزرگی/کوچکیِ حروف (استایلِ «تمام‌حروف‌بزرگ» با ToUpper در متن پخته می‌شود) حذف
        /// می‌شوند. ⚠️ عیناً همان normalize() در tools/enrichment/dump_docx.py.
        /// </summary>
        public static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            string t = text.Replace("{blk}", "").Replace("{/blk}", "");
            t = t.Replace('‘', '\'').Replace('’', '\'')
                 .Replace('“', '"').Replace('”', '"')
                 .Replace('–', '-').Replace('—', '-')
                 .Replace(' ', ' ').Replace("​", "");
            t = Regex.Replace(t, @"\s+", " ").Trim().ToLowerInvariant();
            return t;
        }

        /// <summary>متنِ قابلِ ترجمه‌ی پاراگراف: اسپن‌های متنی پشتِ‌هم (جدول و عکس جدا).</summary>
        public static string ParagraphText(ParagraphData p)
        {
            var sb = new StringBuilder();
            foreach (var s in p.Spans)
            {
                if (s.Type == "text" && s.Content != null) sb.Append(s.Content);
            }
            return sb.ToString();
        }

        private static bool HasLetters(string text) =>
            !string.IsNullOrEmpty(text) && Regex.IsMatch(text, @"[A-Za-z]");

        // ───────────────────────── اعمال ─────────────────────────

        public static Result Apply(string bookDir, List<PageData> pages, List<AudioScriptTrack> audioScripts)
        {
            var result = new Result();
            var tm = LoadTranslationMemory(bookDir);
            var lexicon = LoadLexicon(bookDir);

            var usedTerms = new HashSet<string>(StringComparer.Ordinal);
            var pending = new List<JObject>();
            var pendingKeys = new HashSet<string>(StringComparer.Ordinal);
            var bookText = new StringBuilder();

            void AddPending(int page, string unit, string section, string where, string key, string text)
            {
                // هر متن فقط یک‌بار در pending (متنِ تکراری یک ترجمه دارد)
                if (!pendingKeys.Add(key)) return;
                pending.Add(new JObject
                {
                    ["page"] = page,
                    ["unit"] = unit,
                    ["section"] = section,
                    ["where"] = where,
                    ["key"] = key,
                    ["text"] = text,
                });
            }

            void Visit(ParagraphData p, int page, string unit, string section, string where)
            {
                string text = ParagraphText(p);
                SpanData merged = MergedBlankSpan(p);
                if (merged != null && merged.HiddenParagraphs != null && merged.HiddenParagraphs.Count > 0)
                {
                    // G-7: جوابِ ادغام‌شده‌ی BlkPa با جدول/عکس. هر پاراگرافِ اصلیِ جواب (و
                    // سلول‌های جدولش) مثلِ سند جدا ترجمه می‌شود؛ ترجمه‌ی کلِ جواب = ترجمه‌های
                    // متنی‌ها پشتِ هم.
                    foreach (var hp in merged.HiddenParagraphs)
                        Visit(hp, page, unit, section, where);
                    var withText = merged.HiddenParagraphs.Where(h => HasLetters(ParagraphText(h))).ToList();
                    p.TranslationFa = JoinAll(withText.Select(h => h.TranslationFa));
                    p.TranslationAr = JoinAll(withText.Select(h => h.TranslationAr));
                }
                else if (HasLetters(text))
                {
                    result.TotalParagraphs++;
                    bookText.Append(text.Replace("{blk}", "").Replace("{/blk}", "")).Append('\n');
                    string key = Normalize(text);
                    if (tm.TryGetValue(key, out var e))
                    {
                        p.TranslationFa = string.IsNullOrWhiteSpace(e.Fa) ? null : e.Fa;
                        p.TranslationAr = string.IsNullOrWhiteSpace(e.Ar) ? null : e.Ar;
                        foreach (var t in e.Terms) usedTerms.Add(t);
                        result.Translated++;
                    }
                    else if (merged != null)
                    {
                        // G-7: چند پاراگرافِ BlkPa که اکسترکتور در یک جای‌خالی ادغام کرده؛ در
                        // سند (و در dump_docx) هر کدام پاراگرافِ جدایی است، پس هر خط جدا.
                        var fa = new List<string>();
                        var ar = new List<string>();
                        bool all = true;
                        foreach (string line in MergedBlankLines(merged))
                        {
                            if (!HasLetters(line)) continue;
                            var le = LookupLine(tm, line, out string lineKey, out string lineText);
                            if (le == null)
                            {
                                all = false;
                                AddPending(page, unit, section, where, lineKey, lineText);
                                continue;
                            }
                            fa.Add(le.Fa ?? "");
                            ar.Add(le.Ar ?? "");
                            foreach (var t in le.Terms) usedTerms.Add(t);
                        }
                        if (all)
                        {
                            p.TranslationFa = JoinAll(fa);
                            p.TranslationAr = JoinAll(ar);
                            result.Translated++;
                        }
                        else
                        {
                            p.TranslationFa = null;
                            p.TranslationAr = null;
                            result.Pending++;
                        }
                    }
                    else
                    {
                        p.TranslationFa = null;
                        p.TranslationAr = null;
                        result.Pending++;
                        AddPending(page, unit, section, where, key, text);
                    }
                }
                // سلول‌های جدول: هر پاراگرافِ سلول جداگانه
                foreach (var s in p.Spans)
                {
                    if (s.TableRows == null) continue;
                    foreach (var row in s.TableRows)
                        foreach (var cell in row.Cells)
                            foreach (var cp in cell.Paragraphs)
                                Visit(cp, page, unit, section, "table");
                }
            }

            string curUnit = null, curSection = null;
            foreach (var page in pages)
            {
                foreach (var p in page.Paragraphs)
                {
                    TrackOutline(p, ref curUnit, ref curSection);
                    Visit(p, page.PageNumber, curUnit, curSection, "page");
                }
            }
            if (audioScripts != null)
            {
                foreach (var track in audioScripts)
                    foreach (var p in track.Paragraphs)
                        Visit(p, 0, "AUDIO SCRIPT", track.AudioTrackName, "audio");
            }

            result.Interactives = BuildInteractives(lexicon, usedTerms, bookText.ToString());
            result.InteractiveEntries = result.Interactives?.Count ?? 0;

            // گزارشِ بدونِ‌ترجمه‌ها (و پاک‌کردنِ گزارشِ کهنه اگر همه ترجمه دارند)
            string dir = Path.Combine(bookDir, EnrichmentFolder);
            result.PendingPath = Path.Combine(dir, PendingFile);
            if (pending.Count > 0)
            {
                Directory.CreateDirectory(dir);
                var report = new JObject
                {
                    ["generatedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                    ["total"] = result.TotalParagraphs,
                    ["translated"] = result.Translated,
                    ["pendingUnique"] = pending.Count,
                    ["items"] = new JArray(pending),
                };
                File.WriteAllText(result.PendingPath, report.ToString(Formatting.Indented), new UTF8Encoding(false));
            }
            else if (File.Exists(result.PendingPath))
            {
                File.Delete(result.PendingPath);
            }
            return result;
        }

        // ───────────────────────── جواب‌های ادغام‌شده (G-7) ─────────────────────────

        /// <summary>
        /// اسپنِ جای‌خالی‌ای که اکسترکتور از ادغامِ پاراگراف‌های BlkPa ساخته (MainForm،
        /// MergeBlankParagraphs): پاراگراف فقط همین یک اسپنِ متنی را دارد، متنش سراسر
        /// {blk}…{/blk} است، و یا چند خط (InnerSpanِ «\n») دارد یا محتوای غنی.
        /// </summary>
        private static SpanData MergedBlankSpan(ParagraphData p)
        {
            var texts = p.Spans.Where(s => s.Type == "text" && !string.IsNullOrEmpty(s.Content)).ToList();
            if (texts.Count != 1) return null;
            var s0 = texts[0];
            if (!s0.Content.StartsWith("{blk}") || !s0.Content.EndsWith("{/blk}")) return null;
            bool multiLine = s0.InnerSpans != null && s0.InnerSpans.Any(x => x.Content == "\n");
            bool rich = s0.HiddenParagraphs != null && s0.HiddenParagraphs.Count > 0;
            return multiLine || rich ? s0 : null;
        }

        /// <summary>خط‌های جوابِ ادغام‌شده (هر خط = یک پاراگرافِ اصلیِ سند).</summary>
        private static List<string> MergedBlankLines(SpanData blank)
        {
            var lines = new List<string>();
            var sb = new StringBuilder();
            foreach (var s in blank.InnerSpans ?? new List<SpanData>())
            {
                if (s.Content == "\n") { lines.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(s.Content);
            }
            lines.Add(sb.ToString());
            return lines;
        }

        /// <summary>
        /// شماره‌ی لیستی که اکسترکتور اولِ خط‌های دوم به بعد گذاشته («2:  »، «b)  »، «•  »).
        /// در سند شماره جزءِ متن نیست، پس کلیدِ dump_docx آن را ندارد.
        /// </summary>
        private static readonly Regex InjectedListMarker = new Regex(
            @"^\s*(?:\(?(?:[0-9]{1,3}|[a-zA-Z]|[ivxIVX]{1,5})[.:)]|[•●▪■◦○\-–])\s+",
            RegexOptions.Compiled);

        private static TmEntry LookupLine(Dictionary<string, TmEntry> tm, string line, out string key, out string text)
        {
            text = line;
            key = Normalize(line);
            if (tm.TryGetValue(key, out var e)) return e;
            string stripped = InjectedListMarker.Replace(line, "", 1);
            if (stripped != line)
            {
                text = stripped;
                key = Normalize(stripped);
                if (tm.TryGetValue(key, out e)) return e;
            }
            return null;
        }

        /// <summary>همه با \n پشتِ هم؛ اگر یکی خالی است null (ترجمه‌ی نصفه نشان داده نمی‌شود).</summary>
        private static string JoinAll(IEnumerable<string> parts)
        {
            var list = parts.ToList();
            if (list.Count == 0 || list.Any(string.IsNullOrWhiteSpace)) return null;
            return string.Join("\n", list.Select(x => x.Trim()));
        }

        private static void TrackOutline(ParagraphData p, ref string unit, ref string section)
        {
            if (p.OutlineLevel == null) return;
            string title = ParagraphText(p).Replace("{blk}", "").Replace("{/blk}", "").Trim();
            if (title.Length == 0) return;
            if (p.OutlineLevel == 1) { unit = title; section = null; }
            else if (p.OutlineLevel == 2) section = title;
        }

        // ───────────────────────── حافظه‌ی ترجمه ─────────────────────────

        /// <summary>
        /// G-2: اول ترجمه‌های همین کتاب، بعد (برای متن‌هایی که این کتاب ندارد) ترجمه‌های
        /// کتاب‌های هم‌سطح در همان پوشه‌ی داده — جمله‌های تکراریِ دستورِ تمرین‌ها یک‌بار ترجمه
        /// می‌شوند.
        /// </summary>
        private static Dictionary<string, TmEntry> LoadTranslationMemory(string bookDir)
        {
            var tm = new Dictionary<string, TmEntry>(StringComparer.Ordinal);
            var files = new List<string> { Path.Combine(bookDir, EnrichmentFolder, TranslationsFile) };
            var parent = Directory.GetParent(Path.GetFullPath(bookDir));
            if (parent != null)
            {
                foreach (var sibling in Directory.GetDirectories(parent.FullName).OrderBy(d => d, StringComparer.Ordinal))
                {
                    string f = Path.Combine(sibling, EnrichmentFolder, TranslationsFile);
                    if (!files.Any(x => string.Equals(Path.GetFullPath(x), Path.GetFullPath(f), StringComparison.OrdinalIgnoreCase)))
                        files.Add(f);
                }
            }
            foreach (var f in files)
            {
                if (!File.Exists(f)) continue;
                var root = JObject.Parse(File.ReadAllText(f, Encoding.UTF8));
                if (!(root["entries"] is JObject entries)) continue;
                foreach (var kv in entries)
                {
                    string key = Normalize(kv.Key);   // کلیدِ دستی‌ویرایش‌شده هم نرمال شود
                    if (tm.ContainsKey(key) || !(kv.Value is JObject v)) continue;
                    var e = new TmEntry
                    {
                        Fa = (string)v["fa"],
                        Ar = (string)v["ar"],
                    };
                    if (v["terms"] is JArray terms)
                        e.Terms.AddRange(terms.Select(t => (string)t).Where(t => !string.IsNullOrWhiteSpace(t)));
                    tm[key] = e;
                }
            }
            return tm;
        }

        // ───────────────────────── دیکشنریِ مشترک ─────────────────────────

        /// <summary>G-3: _lexicon/lexicon.json را در پوشه‌ی کتاب یا تا سه پوشه بالاتر پیدا می‌کند.</summary>
        private static Dictionary<string, JObject> LoadLexicon(string bookDir)
        {
            var lex = new Dictionary<string, JObject>(StringComparer.Ordinal);
            var dir = new DirectoryInfo(Path.GetFullPath(bookDir));
            for (int i = 0; i < 4 && dir != null; i++, dir = dir.Parent)
            {
                string f = Path.Combine(dir.FullName, LexiconFolder, LexiconFile);
                if (!File.Exists(f)) continue;
                var root = JObject.Parse(File.ReadAllText(f, Encoding.UTF8));
                if (root["entries"] is JObject entries)
                {
                    foreach (var kv in entries)
                        if (kv.Value is JObject v) lex[kv.Key] = v;
                }
                break;
            }
            return lex;
        }

        /// <summary>
        /// G-4: Interactivesِ کتاب = مدخل‌هایی از دیکشنریِ مشترک که ترجمه‌های همین کتاب به
        /// آن‌ها اشاره کرده‌اند (terms). برای هر شکلِ مدخل (forms) هر صورتی که واقعاً در متنِ
        /// کتاب آمده (با همان بزرگی/کوچکیِ حروف؛ چون اپ دقیق تطبیق می‌دهد) یک آیتم می‌شود.
        /// قالبِ آیتم همان InteractiveWordِ اپ است + فیلدهای تازه (ipa، pos، lexiconId) که
        /// نسخه‌ی فعلیِ اپ نادیده می‌گیرد.
        /// </summary>
        private static List<JObject> BuildInteractives(
            Dictionary<string, JObject> lexicon, HashSet<string> usedTerms, string bookText)
        {
            if (lexicon.Count == 0 || usedTerms.Count == 0) return null;
            var list = new List<JObject>();
            var seenSurface = new HashSet<string>(StringComparer.Ordinal);
            foreach (var id in usedTerms.OrderBy(x => x, StringComparer.Ordinal))
            {
                if (!lexicon.TryGetValue(id, out var entry)) continue;
                var forms = (entry["forms"] as JArray)?.Select(f => (string)f).Where(f => !string.IsNullOrWhiteSpace(f)).ToList()
                            ?? new List<string>();
                string head = (string)entry["headword"];
                if (!string.IsNullOrWhiteSpace(head) && !forms.Contains(head)) forms.Insert(0, head);

                foreach (var form in forms)
                {
                    var rx = new Regex(@"(?<![A-Za-z0-9])" + Regex.Escape(form) + @"(?![A-Za-z0-9])",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    foreach (Match m in rx.Matches(bookText))
                    {
                        string surface = m.Value;
                        if (!seenSurface.Add(surface)) continue;   // هر صورت یک‌بار (اولین مدخل برنده)
                        list.Add(new JObject
                        {
                            ["exactText"] = surface,
                            ["translationFa"] = entry["translationFa"],
                            ["translationAr"] = entry["translationAr"],
                            ["pronounceFa"] = entry["pronounceFa"],
                            ["pronounceAr"] = entry["pronounceAr"],
                            ["explanationFa"] = entry["explanationFa"],
                            ["explanationAr"] = entry["explanationAr"],
                            ["cefrLevel"] = entry["cefrLevel"],
                            ["ipa"] = entry["ipa"],
                            ["pos"] = entry["pos"],
                            ["lexiconId"] = id,
                        });
                    }
                }
            }
            // طولانی‌ترها اول (همان ترتیبی که اپ برای تطبیق لازم دارد)
            return list.OrderByDescending(o => ((string)o["exactText"]).Length).ToList();
        }
    }
}
