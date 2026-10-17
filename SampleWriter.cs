using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CustomLayoutGenerator
{
    /// <summary>
    /// 🌟 SM-1 … SM-6 (RULES-extraction.md): ساختنِ «نسخه‌ی نمونه»ِ کتاب، هم‌زمان با استخراج.
    ///
    /// ورودی همان خروجیِ تازه‌نوشته‌شده‌ی BookOutputWriter است (index.json، pages/،
    /// audio_scripts/، images/ و پوشه‌ی audio/ِ کنارِ docx)، نه مدل‌های حافظه — پس صفحه‌های
    /// نمونه بایت‌به‌بایت همان صفحه‌های کتاب‌اند (با همان هَش و همان ترجمه‌ها) و این کلاس
    /// را می‌شود جدا از فرمِ ویندوز هم اجرا و آزمایش کرد.
    ///
    /// خروجی: پوشه‌ی sample/ کنارِ خروجیِ اصلی، با همان ساختار:
    ///   sample/index.json      Pages فقط صفحه‌های یک Unit، ولی Outline کامل (همه‌ی Unitها)
    ///                          + بخشِ Sample (عنوان و بازه‌ی صفحه‌های نمونه و کلِ کتاب)
    ///   sample/pages/          کپیِ صفحه‌های همان Unit
    ///   sample/audio_scripts/  فقط اسکریپتِ فایل‌های صوتی‌ای که در این صفحه‌ها لینک دارند
    ///   sample/images/         فقط عکس‌های همین صفحه‌ها
    ///   sample/audio/          فقط mp3های لازم (از پوشه‌ی audio/ِ کنارِ docx، اگر هست)
    /// اپ با دیدنِ بخشِ Sample، فهرستِ همه‌ی Unitها را نشان می‌دهد و برای Unitهای بیرون از
    /// نمونه پیامِ «خریداری نشده» می‌دهد (قانونِ N-12ِ RULES-rendering.md).
    /// </summary>
    public static class SampleWriter
    {
        public const string FolderName = "sample";

        public class Result
        {
            public bool Written;
            public string Title;
            public int FirstPage, LastPage, PageCount, AudioScriptCount, ImageCount, AudioFileCount, InteractiveCount;
            public List<string> MissingAudio = new List<string>();
            public List<string> MissingImages = new List<string>();
            public string Summary;
        }

        /// <summary>عنوانِ گره‌های سطحِ اولِ Outline (در Mindset: Unitها) از index.json؛ خالی اگر ندارد.</summary>
        public static List<string> TopLevelTitles(string outputDir)
        {
            var index = LoadIndex(outputDir);
            return (index?["Outline"] as JArray)?.Select(e => (string)e["Title"] ?? "").ToList()
                   ?? new List<string>();
        }

        /// <summary>
        /// نسخه‌ی نمونه را از گره‌ی شماره‌ی <paramref name="topIndex"/>ِ سطحِ اولِ Outline می‌سازد.
        /// پوشه‌ی sample/ِ قبلی (فقط index.json و چهار زیرپوشه‌ی شناخته‌شده‌اش) اول پاک می‌شود
        /// تا فایلِ کهنه‌ای از نمونه‌ی قبلی نماند.
        /// </summary>
        public static Result Write(string outputDir, int topIndex)
        {
            var result = new Result();
            var index = LoadIndex(outputDir);
            var outline = index?["Outline"] as JArray;
            var pages = (index?["Pages"] as JArray)?.OfType<JObject>().ToList();
            if (index == null || pages == null || pages.Count == 0)
            {
                result.Summary = "نسخه‌ی نمونه ساخته نشد: index.json یا صفحه‌ها پیدا نشد.";
                return result;
            }
            if (outline == null || outline.Count == 0)
            {
                result.Summary = "نسخه‌ی نمونه ساخته نشد: سند فهرستِ ساختاری (Outline Level) ندارد، پس معلوم نیست Unitها کجا شروع می‌شوند.";
                return result;
            }
            if (topIndex < 0 || topIndex >= outline.Count) topIndex = 0;

            // ── SM-2: بازه‌ی صفحه‌ها ──────────────────────────────────────────────
            var numbers = pages.Select(p => (int)p["N"]).ToList();
            int bookFirst = numbers.Min(), bookLast = numbers.Max();
            var chosen = (JObject)outline[topIndex];
            int first = topIndex == 0 ? bookFirst : (int)chosen["PageNumber"];   // Unitِ اول: جلد و صفحه‌های پیش از آن هم
            int last = bookLast;
            if (topIndex + 1 < outline.Count)
            {
                var next = (JObject)outline[topIndex + 1];
                int nextPage = (int)next["PageNumber"];
                bool nextAtTop = (bool?)next["AtPageTop"] == true || (int?)next["ParaIndex"] == 0;
                // اگر Unitِ بعدی وسطِ صفحه شروع شود، آن صفحه هنوز پایانِ همین Unit را دارد
                last = nextAtTop ? nextPage - 1 : nextPage;
            }
            var samplePages = pages.Where(p => (int)p["N"] >= first && (int)p["N"] <= last).ToList();
            if (samplePages.Count == 0)
            {
                result.Summary = $"نسخه‌ی نمونه ساخته نشد: بازه‌ی صفحه‌های «{(string)chosen["Title"]}» خالی است.";
                return result;
            }

            string sampleDir = Path.Combine(outputDir, FolderName);
            CleanSampleDir(sampleDir);
            Directory.CreateDirectory(Path.Combine(sampleDir, "pages"));

            var text = new StringBuilder();
            var imageNames = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            var audioNames = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

            // ── SM-3: صفحه‌ها (کپیِ بی‌تغییر) ─────────────────────────────────────
            foreach (var entry in samplePages)
            {
                string rel = (string)entry["File"];
                string src = Path.Combine(outputDir, rel);
                string dst = Path.Combine(sampleDir, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(dst));
                File.Copy(src, dst, true);
                Collect(JToken.Parse(File.ReadAllText(src)), text, imageNames, audioNames);
            }

            // ── لینک‌های صوتی و اسکریپت‌های صوتیِ همین صفحه‌ها ────────────────────
            var links = new JArray((index["AudioLinksIndex"] as JArray ?? new JArray()).OfType<JObject>()
                .Where(l => (int)l["PageNumber"] >= first && (int)l["PageNumber"] <= last));
            foreach (var l in links) audioNames.Add((string)l["FileName"] ?? "");
            audioNames.Remove("");

            var tracks = new JArray();
            foreach (var t in (index["AudioScripts"] as JArray ?? new JArray()).OfType<JObject>())
            {
                if (!audioNames.Contains((string)t["AudioTrackName"] ?? "")) continue;
                string rel = (string)t["File"];
                string src = Path.Combine(outputDir, rel ?? "");
                if (string.IsNullOrEmpty(rel) || !File.Exists(src)) continue;
                string dst = Path.Combine(sampleDir, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(dst));
                File.Copy(src, dst, true);
                Collect(JToken.Parse(File.ReadAllText(src)), text, imageNames, audioNames: null);
                tracks.Add(t.DeepClone());
            }

            // ── عکس‌ها و فایل‌های صوتی ────────────────────────────────────────────
            foreach (var img in imageNames)
            {
                string src = Path.Combine(outputDir, "images", img);
                if (!File.Exists(src)) { result.MissingImages.Add(img); continue; }
                Directory.CreateDirectory(Path.Combine(sampleDir, "images"));
                File.Copy(src, Path.Combine(sampleDir, "images", img), true);
                result.ImageCount++;
            }
            foreach (var mp3 in audioNames)
            {
                string src = Path.Combine(outputDir, "audio", mp3);
                if (!File.Exists(src)) { result.MissingAudio.Add(mp3); continue; }
                Directory.CreateDirectory(Path.Combine(sampleDir, "audio"));
                File.Copy(src, Path.Combine(sampleDir, "audio", mp3), true);
                result.AudioFileCount++;
            }

            // ── SM-4: کلماتِ تعاملی فقط آن‌هایی که در متنِ نمونه هستند ──────────────
            string sampleText = text.ToString();
            var interactives = new JArray();
            foreach (var w in (index["Interactives"] as JArray ?? new JArray()).OfType<JObject>())
            {
                string surface = (string)w["exactText"];
                if (string.IsNullOrEmpty(surface)) continue;
                // اپ دقیق (با همان بزرگی/کوچکیِ حروف) تطبیق می‌دهد؛ این‌جا هم همان‌طور
                if (Regex.IsMatch(sampleText, @"(?<![A-Za-z0-9])" + Regex.Escape(surface) + @"(?![A-Za-z0-9])"))
                    interactives.Add(w.DeepClone());
            }

            // ── SM-5: index.jsonِ نمونه ────────────────────────────────────────────
            var sampleIndex = new JObject
            {
                ["SchemaVersion"] = index["SchemaVersion"] ?? 2,
                ["BookVersion"] = ShortHash(string.Concat(samplePages.Select(p => (string)p["Version"]))),
                ["Pages"] = new JArray(samplePages.Select(p => p.DeepClone())),
            };
            if (interactives.Count > 0) sampleIndex["Interactives"] = interactives;
            sampleIndex["AudioScripts"] = tracks;
            sampleIndex["AudioLinksIndex"] = links;
            sampleIndex["Outline"] = outline.DeepClone();   // کامل: اپ همه‌ی Unitها را نشان می‌دهد
            sampleIndex["Sample"] = new JObject
            {
                ["Title"] = (string)chosen["Title"],
                ["FirstPage"] = first,
                ["LastPage"] = last,
                ["BookFirstPage"] = bookFirst,
                ["BookLastPage"] = bookLast,
            };
            File.WriteAllText(Path.Combine(sampleDir, "index.json"),
                sampleIndex.ToString(Formatting.Indented), new UTF8Encoding(false));

            result.Written = true;
            result.Title = (string)chosen["Title"];
            result.FirstPage = first;
            result.LastPage = last;
            result.PageCount = samplePages.Count;
            result.AudioScriptCount = tracks.Count;
            result.InteractiveCount = interactives.Count;

            var sb = new StringBuilder();
            sb.Append($"نسخه‌ی نمونه: «{result.Title}» — صفحه‌های {first} تا {last} ({result.PageCount} صفحه)، ");
            sb.Append($"{result.AudioScriptCount} اسکریپتِ صوتی، {result.ImageCount} عکس، {result.AudioFileCount} فایلِ صوتی، ");
            sb.Append($"{result.InteractiveCount} کلمه‌ی تعاملی ← پوشه‌ی {FolderName}");
            if (result.MissingAudio.Count > 0)
                sb.Append($"\n⚠️ {result.MissingAudio.Count} فایلِ صوتیِ نمونه در پوشه‌ی audio/ِ کنارِ docx پیدا نشد (بعداً دستی در sample\\audio بگذارید): "
                          + string.Join("، ", result.MissingAudio.Take(8)) + (result.MissingAudio.Count > 8 ? "، …" : ""));
            if (result.MissingImages.Count > 0)
                sb.Append($"\n⚠️ {result.MissingImages.Count} عکسِ نمونه در images/ پیدا نشد.");
            result.Summary = sb.ToString();
            return result;
        }

        private static JObject LoadIndex(string outputDir)
        {
            string path = Path.Combine(outputDir, "index.json");
            if (!File.Exists(path)) return null;
            try { return JObject.Parse(File.ReadAllText(path)); }
            catch { return null; }
        }

        /// <summary>فقط چیزهایی که خودِ این کلاس می‌سازد پاک می‌شوند، نه هر چیزی در sample/.</summary>
        private static void CleanSampleDir(string sampleDir)
        {
            if (!Directory.Exists(sampleDir)) return;
            string idx = Path.Combine(sampleDir, "index.json");
            if (File.Exists(idx)) File.Delete(idx);
            foreach (var sub in new[] { "pages", "audio_scripts", "images", "audio" })
            {
                string d = Path.Combine(sampleDir, sub);
                if (Directory.Exists(d)) Directory.Delete(d, true);
            }
        }

        /// <summary>
        /// گشتِ بازگشتی در JSONِ صفحه (شاملِ جدول‌ها، چیدمان‌ها و محتوای پنهانِ جواب‌ها): متنِ
        /// همه‌ی Contentها، نامِ عکس‌ها (Type=image ← Url) و فایل‌های صوتیِ لینک‌شده (Url=audio:…).
        /// روی JSON کار می‌کند نه مدل‌ها، تا هر ساختارِ تودرتوی تازه‌ای هم خودبه‌خود پوشش داده شود.
        /// </summary>
        private static void Collect(JToken token, StringBuilder text, ISet<string> images, ISet<string> audioNames)
        {
            if (token is JObject o)
            {
                string type = (string)o["Type"];
                string url = o["Url"]?.Type == JTokenType.String ? (string)o["Url"] : null;
                if (type == "image" && !string.IsNullOrEmpty(url)) images.Add(Path.GetFileName(url));
                else if (url != null && url.StartsWith("audio:") && audioNames != null)
                    audioNames.Add(url.Substring("audio:".Length));
                // اسپن‌های یک پاراگراف پشتِ‌هم (یک عبارت ممکن است در چند اسپن شکسته باشد)،
                // و یک فاصله بعد از هر پاراگراف
                if (o["Content"]?.Type == JTokenType.String) text.Append((string)o["Content"]);
                foreach (var p in o.Properties()) Collect(p.Value, text, images, audioNames);
                if (o["Spans"] is JArray) text.Append(' ');
            }
            else if (token is JArray a)
            {
                foreach (var t in a) Collect(t, text, images, audioNames);
            }
        }

        private static string ShortHash(string content)
        {
            using (var sha = SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(content));
                var sb = new StringBuilder(16);
                for (int i = 0; i < 8; i++) sb.Append(hash[i].ToString("x2"));
                return sb.ToString();
            }
        }
    }
}
