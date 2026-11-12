using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WinForms = System.Windows.Forms; // با alias تا تداخل نام با Open XML پیش نیاید

namespace CustomLayoutGenerator
{
    internal class SpeakerMerger
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

        public static int Process(WordprocessingDocument doc, string styleName)
        {
            var main = doc.MainDocumentPart!;
            var styles = main.StyleDefinitionsPart?.Styles
                ?? throw new Exception("فایل بخش استایل ندارد.");

            // پیدا کردن استایل از روی نام (یا شناسه)
            var speakerStyle = styles.Elements<Style>().FirstOrDefault(s =>
                s.Type?.Value == StyleValues.Paragraph &&
                (string.Equals(s.StyleName?.Val?.Value, styleName, StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(s.StyleId?.Value, styleName, StringComparison.OrdinalIgnoreCase)))
                ?? throw new Exception($"استایل پاراگرافی با نام «{styleName}» در این فایل پیدا نشد.");

            string styleId = speakerStyle.StyleId!.Value!;

            // ظاهر نهایی استایل (با در نظر گرفتن BasedOn) برای نوشتن مستقیم روی Runها
            var styleRunProps = BuildStyleRunProps(styles, speakerStyle);

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

                MergeInto(p, next, styleRunProps);
                merged++;
            }

            main.Document.Save();
            return merged;
        }

        static bool HasStyle(Paragraph p, string styleId) =>
            p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == styleId;

        static void MergeInto(Paragraph speaker, Paragraph next, RunProperties styleRunProps)
        {
            // ۱) ظاهر استایل NotSpoken را مستقیم روی Runهای گوینده می‌نویسیم
            foreach (var run in speaker.Descendants<Run>().ToList())
            {
                var props = (RunProperties)styleRunProps.CloneNode(true);
                if (run.RunProperties != null)
                {
                    Overlay(props, run.RunProperties); // فرمت خودِ Run اولویت دارد
                    run.RunProperties.Remove();
                }
                SortRunProps(props);
                run.PrependChild(props);
            }

            // ۲) فاصله‌های انتهای متن گوینده را پاک می‌کنیم
            var texts = speaker.Descendants<Text>().ToList();
            for (int i = texts.Count - 1; i >= 0; i--)
            {
                if (string.IsNullOrWhiteSpace(texts[i].Text)) { texts[i].Text = ""; continue; }
                texts[i].Text = texts[i].Text.TrimEnd();
                texts[i].Space = SpaceProcessingModeValues.Preserve;
                break;
            }

            // ۳) اضافه کردن ": " با ظاهر گوینده
            var colonProps = (RunProperties)styleRunProps.CloneNode(true);
            SortRunProps(colonProps);
            var colonRun = new Run(colonProps);
            colonRun.Append(new Text(": ") { Space = SpaceProcessingModeValues.Preserve });
            speaker.Append(colonRun);

            // ۴) تنظیمات پاراگراف (از جمله استایل) از پاراگراف دوم گرفته می‌شود
            speaker.ParagraphProperties?.Remove();
            if (next.ParagraphProperties != null)
                speaker.PrependChild((ParagraphProperties)next.ParagraphProperties.CloneNode(true));

            // ۵) انتقال محتوای پاراگراف دوم (Run، Hyperlink و ...) و حذف آن
            foreach (var child in next.ChildElements.Where(c => c is not ParagraphProperties).ToList())
            {
                child.Remove();
                speaker.Append(child);
            }
            next.Remove();
        }

        // ساخت rPr نهایی استایل با دنبال کردن زنجیره‌ی BasedOn
        static RunProperties BuildStyleRunProps(Styles styles, Style style)
        {
            var chain = new List<Style>();
            Style? cur = style;
            while (cur != null && chain.Count < 20)
            {
                chain.Insert(0, cur); // پایه اول، مشتق‌شده آخر
                string? baseId = cur.BasedOn?.Val?.Value;
                cur = baseId == null ? null
                    : styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == baseId);
            }

            var result = new RunProperties();
            foreach (var s in chain)
                if (s.StyleRunProperties != null)
                    Overlay(result, s.StyleRunProperties);

            SortRunProps(result);
            return result;
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
        static void SortRunProps(RunProperties rp)
        {
            var sorted = rp.ChildElements
                .OrderBy(e => { int i = Array.IndexOf(RPrOrder, e.LocalName); return i < 0 ? int.MaxValue : i; })
                .Select(e => e.CloneNode(true))
                .ToList();
            rp.RemoveAllChildren();
            foreach (var e in sorted) rp.AppendChild(e);
        }
    }
}
