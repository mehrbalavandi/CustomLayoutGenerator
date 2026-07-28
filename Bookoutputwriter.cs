using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace WordToJsonParser
{
    /// <summary>
    /// خروجی را به‌جای یک فایلِ واحد، به‌صورت «یک index.json + یک فایل به‌ازای هر صفحه»
    /// می‌نویسد. نسخهٔ هر صفحه = هَشِ محتوای همان صفحه؛ بنابراین در بازتولیدِ کل کتاب،
    /// فقط صفحاتی که واقعاً تغییر کرده‌اند هَش جدید می‌گیرند و کلاینت دقیقاً همان‌ها را
    /// دوباره دانلود می‌کند.
    ///
    /// ⚠️ مهم: چون مرحلهٔ هوش مصنوعی (ترجمه/Interactives) محتوای صفحات را عوض می‌کند،
    /// این متد باید «آخرین» گامِ خط‌لوله باشد — یعنی بعد از تزریقِ ترجمه‌ها — تا هَش‌ها
    /// محتوای نهاییِ ارسال‌شده را بازتاب دهند.
    /// </summary>
    public static class BookOutputWriter
    {
        private static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            NullValueHandling = NullValueHandling.Ignore
            // برای فایل کوچک‌تر، ContractResolver = new OmitEmptyContractResolver()
            // را از تغییراتِ «کاهش حجم» قبلی هم می‌توانید اینجا اضافه کنید.
        };

        public static void Write(
            string outputDir,
            List<PageData> pages,
            List<AudioScriptTrack> audioScripts,
            object interactives = null)   // دیکشنری Interactives سطح کتاب (توسط مرحلهٔ AI پر می‌شود)
        {
            var pagesDir = Path.Combine(outputDir, "pages");
            Directory.CreateDirectory(pagesDir);

            var manifest = new List<PageIndexEntry>();

            for (int i = 0; i < pages.Count; i++)
            {
                var page = pages[i];
                int n = page.PageNumber > 0 ? page.PageNumber : (i + 1);
                string fileName = $"page_{n:D4}.json";

                // فرمت None (بدون فاصله) تا هَش پایدار و مستقل از قالب‌بندی باشد
                string pageJson = JsonConvert.SerializeObject(page, Formatting.None, Settings);
                File.WriteAllText(Path.Combine(pagesDir, fileName), pageJson);

                manifest.Add(new PageIndexEntry
                {
                    N = n,
                    File = $"pages/{fileName}",
                    Version = ShortHash(pageJson)   // 🌟 نسخه = هَشِ محتوای صفحه
                });
            }

            var index = new BookIndex
            {
                SchemaVersion = 2,
                // اگر هر صفحه‌ای عوض شود این هم عوض می‌شود → پرچمِ ارزانِ «چیزی جدید هست؟»
                BookVersion = ShortHash(string.Concat(manifest.Select(m => m.Version))),
                Pages = manifest,
                Interactives = interactives,
                AudioScripts = audioScripts,
                // 🐞 شاخصِ سطحِ‌کتابِ لینک‌های صوتیِ داخلِ متن — دقیقاً همان
                // چیزی که سمتِ فلاتر (buildBookAudioPlaylist) قبلاً مجبور
                // بود با گشتنِ زنده در محتوای *همه‌ی* صفحاتِ لودشده بسازد؛
                // چون این کار نیازِ به لودِ کاملِ کتاب دارد، مانعِ اصلیِ
                // لودِ تنبل/صفحه‌به‌صفحه بود. حالا از قبل، همین‌جا در زمانِ
                // استخراج، محاسبه و در index.json نوشته می‌شود.
                AudioLinksIndex = BuildAudioLinksIndex(pages)
            };

            string indexJson = JsonConvert.SerializeObject(index, Formatting.Indented, Settings);
            File.WriteAllText(Path.Combine(outputDir, "index.json"), indexJson);
        }

        // 🐞 اسکنِ یک‌بارِ کلِ کتاب (شاملِ پاراگراف‌های داخلِ سلول‌های جدول،
        // به‌صورتِ بازگشتی) برای هر اسپنی که Url اش با "audio:" شروع می‌شود —
        // همان قراردادی که هایپرلینک‌های صوتیِ داخلِ متن با آن مشخص می‌شوند.
        // موقعیتِ گزارش‌شده همیشه ParaIndexِ پاراگرافِ *بیرونی* است (نه اندیسِ
        // داخلیِ سلول)، چون این دقیقاً همان قراردادی است که سمتِ فلاتر برای
        // «برو به متن» استفاده می‌کند.
        private static List<AudioLinkEntry> BuildAudioLinksIndex(List<PageData> pages)
        {
            var result = new List<AudioLinkEntry>();

            void ScanSpans(List<SpanData> spans, int pageNumber, int topParaIndex)
            {
                if (spans == null) return;
                foreach (var s in spans)
                {
                    if (!string.IsNullOrEmpty(s.Url) && s.Url.StartsWith("audio:"))
                    {
                        string fileName = s.Url.Substring("audio:".Length);
                        if (!string.IsNullOrEmpty(fileName))
                        {
                            result.Add(new AudioLinkEntry
                            {
                                PageNumber = pageNumber,
                                ParaIndex = topParaIndex,
                                FileName = fileName
                            });
                        }
                    }
                    if (s.Type == "table" && s.TableRows != null)
                    {
                        foreach (var row in s.TableRows)
                        {
                            foreach (var cell in row.Cells)
                            {
                                foreach (var p in cell.Paragraphs)
                                {
                                    ScanSpans(p.Spans, pageNumber, topParaIndex);
                                }
                            }
                        }
                    }
                }
            }

            foreach (var page in pages)
            {
                for (int i = 0; i < page.Paragraphs.Count; i++)
                {
                    ScanSpans(page.Paragraphs[i].Spans, page.PageNumber, i);
                }
            }

            return result;
        }

        private static string ShortHash(string content)
        {
            using (var sha = SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(content));
                var sb = new StringBuilder(16);
                for (int i = 0; i < 8; i++) sb.Append(hash[i].ToString("x2")); // ۱۶ کاراکترِ اول
                return sb.ToString();
            }
        }
    }

    public class BookIndex
    {
        public int SchemaVersion { get; set; }
        public string BookVersion { get; set; }
        public List<PageIndexEntry> Pages { get; set; } = new List<PageIndexEntry>();
        public object Interactives { get; set; }                 // null اگر هنوز تولید نشده
        public List<AudioScriptTrack> AudioScripts { get; set; }  // سطح کتاب — یک آیتم به‌ازای هر فایلِ صوتی
        public List<AudioLinkEntry> AudioLinksIndex { get; set; }  // سطح کتاب — کجای متن دکمه‌ی صوتی هست
    }

    public class PageIndexEntry
    {
        public int N { get; set; }         // شمارهٔ صفحه
        public string File { get; set; }   // مسیر نسبی: pages/page_0001.json
        public string Version { get; set; } // هَشِ محتوا
    }
}